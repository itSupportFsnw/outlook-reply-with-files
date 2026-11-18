dlog("commands.js loaded");

// Ensures event.completed() is called exactly once, with a safety timeout.
function completer(event, label, timeoutMs = 25000) {
  let done = false;
  const finish = (why) => {
    if (done) return;
    done = true;
    dlog(label, "completed:", why);
    event.completed();
  };
  setTimeout(() => finish("timeout"), timeoutMs);
  return finish;
}

// Ribbon / action-bar button: reply with attachments in one click (no task pane).
async function replyWithFiles(event) {
  dlog("replyWithFiles: start");
  const finish = completer(event, "replyWithFiles");
  const item = Office.context.mailbox.item;
  try {
    const { skipped } = await stashAndReply(item);
    if (skipped.length) {
      item.notificationMessages.addAsync("skipped", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `Cloud attachments not copied: ${skipped.join(", ")}`.slice(0, 150),
        icon: "Icon.16x16",
        persistent: false,
      });
    }
    finish("ok");
  } catch (e) {
    dlog("replyWithFiles: error", e && (e.message || e.name || String(e)));
    item.notificationMessages.addAsync("error", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: ("Reply w/ Files failed: " + (e.message || e)).slice(0, 150),
    });
    finish("error");
  }
}

// Event-based handler: when a compose window opens shortly after "Reply w/ Files",
// attach the stashed files to it.
async function onNewMessageComposeHandler(event) {
  dlog("onNewMessageCompose: start");
  const finish = completer(event, "onNewMessageCompose");
  try {
    const n = await attachStashed(Office.context.mailbox.item);
    dlog("onNewMessageCompose: attached", n);
  } catch (e) {
    dlog("onNewMessageCompose: error", e && (e.message || e.name || String(e)));
  }
  finish("done");
}

Office.actions.associate("replyWithFiles", replyWithFiles);
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);

// Required: signals to Outlook that the function file is initialized,
// otherwise commands are never dispatched and the "working on…" bar hangs.
Office.onReady((info) => dlog("Office.onReady", info));
