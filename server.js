// Minimal HTTPS static server for local add-in development (https://localhost:3000).
const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;
const ROOT = __dirname;
const SERVED = new Set([
  "taskpane.html", "commands.html", "commands.js", "shared.js",
  "icon-16.png", "icon-32.png", "icon-80.png", "manifest.xml", "manifest.local.xml",
]);
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".png": "image/png", ".xml": "application/xml" };

(async () => {
  const options = await devCerts.getHttpsServerOptions();
  https.createServer(options, (req, res) => {
    const name = decodeURIComponent(new URL(req.url, "https://localhost").pathname).replace(/^\/+/, "");
    if (name === "log" && req.method === "POST") {
      let body = "";
      req.on("data", (c) => (body += c));
      req.on("end", () => { console.log(new Date().toLocaleTimeString(), "[addin]", body); res.writeHead(204); res.end(); });
      return;
    }
    console.log(new Date().toLocaleTimeString(), req.method, req.url);
    if (!SERVED.has(name)) { res.writeHead(404); return res.end("Not found"); }
    fs.readFile(path.join(ROOT, name), (err, data) => {
      if (err) { res.writeHead(500); return res.end(String(err)); }
      res.writeHead(200, { "Content-Type": TYPES[path.extname(name)] || "application/octet-stream", "Cache-Control": "no-store" });
      res.end(data);
    });
  }).listen(PORT, () => console.log(`Add-in server running at https://localhost:${PORT}`));
})();
