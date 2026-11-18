// Shared between the task pane and the event-based runtime.
// Attachments are stashed in IndexedDB (same origin => shared across both runtimes).

// Debug: send progress lines to the local server console.
function dlog(...args) {
  const msg = args.map((a) => (typeof a === "string" ? a : JSON.stringify(a))).join(" ");
  console.log(msg);
  if (location.hostname !== "localhost") return; // remote log only to the local dev server
  try { fetch("/log", { method: "POST", body: msg, keepalive: true }); } catch (e) {}
}

const STASH_DB = "replyWithAttachments";
const STASH_KEY = "pending";
const STASH_MAX_AGE_MS = 2 * 60 * 1000; // only auto-attach to a compose window opened within 2 min

function openStashDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(STASH_DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore("stash");
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function stashPut(value) {
  const db = await openStashDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("stash", "readwrite");
    tx.objectStore("stash").put(value, STASH_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

// Reads and removes the stash in one transaction so files are only attached once.
async function stashTake() {
  const db = await openStashDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("stash", "readwrite");
    const store = tx.objectStore("stash");
    const getReq = store.get(STASH_KEY);
    getReq.onsuccess = () => store.delete(STASH_KEY);
    tx.oncomplete = () => resolve(getReq.result || null);
    tx.onerror = () => reject(tx.error);
  });
}

function toBase64Utf8(text) {
  const bytes = new TextEncoder().encode(text);
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(bin);
}

function getAttachmentContent(item, id) {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(id, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
      else reject(r.error);
    });
  });
}

// Read mode: copy the message's attachments into the stash, then open the reply form.
// Returns { saved, skipped } where skipped lists cloud attachments that can't be copied.
async function stashAndReply(item) {
  const attachments = (item.attachments || []).filter((a) => !a.isInline);
  dlog("stashAndReply: attachments", attachments.map((a) => `${a.name} (${a.size})`));
  const F = Office.MailboxEnums.AttachmentContentFormat;
  const files = [];
  const skipped = [];

  for (const att of attachments) {
    dlog("fetching", att.name);
    const c = await getAttachmentContent(item, att.id);
    dlog("fetched", att.name, c.format, (c.content || "").length);
    if (c.format === F.Base64) {
      files.push({ name: att.name, base64: c.content });
    } else if (c.format === F.Eml) {
      files.push({ name: att.name.endsWith(".eml") ? att.name : att.name + ".eml", base64: toBase64Utf8(c.content) });
    } else if (c.format === F.ICalendar) {
      files.push({ name: att.name.endsWith(".ics") ? att.name : att.name + ".ics", base64: toBase64Utf8(c.content) });
    } else {
      skipped.push(att.name); // cloud attachments (format "url") can't be copied
    }
  }

  if (files.length) {
    dlog("stashing", files.length);
    await stashPut({ ts: Date.now(), files });
    dlog("stashed");
  }
  dlog("displayReplyForm");
  item.displayReplyForm("");
  dlog("displayReplyForm returned");
  return { saved: files.length, skipped };
}

// Attaches stashed files to a compose item. Returns the number attached.
async function attachStashed(item, { ignoreAge = false } = {}) {
  dlog("attachStashed: start");
  const stash = await stashTake();
  dlog("attachStashed: stash", stash ? { ageMs: Date.now() - stash.ts, files: (stash.files || []).length } : null);
  if (!stash || !stash.files || stash.files.length === 0) return 0;
  if (!ignoreAge && Date.now() - stash.ts > STASH_MAX_AGE_MS) return 0;

  let count = 0;
  for (const file of stash.files) {
    await new Promise((resolve, reject) => {
      item.addFileAttachmentFromBase64Async(file.base64, file.name, { isInline: false }, (r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(r.error);
      });
    });
    count++;
  }
  return count;
}
