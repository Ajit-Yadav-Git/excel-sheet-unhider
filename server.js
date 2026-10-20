// Minimal HTTPS static server for local add-in development (port 3000).
const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const ROOT = path.join(__dirname, "src");
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".py": "text/plain", ".css": "text/css", ".png": "image/png" };

(async () => {
  const opts = await devCerts.getHttpsServerOptions();
  https
    .createServer(opts, (req, res) => {
      const rel = decodeURIComponent(req.url.split("?")[0]);
      const file = path.normalize(path.join(ROOT, rel === "/" ? "taskpane.html" : rel));
      if (!file.startsWith(ROOT)) { res.writeHead(403).end(); return; } // block path traversal
      fs.readFile(file, (err, data) => {
        if (err) { res.writeHead(404).end("Not found"); return; }
        res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream" });
        res.end(data);
      });
    })
    .listen(3000, () => console.log("Serving https://localhost:3000"));
})().catch((e) => { console.error("Failed to start server:", e); process.exit(1); });
