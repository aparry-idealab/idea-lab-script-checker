/**
 * Minimal local HTTPS static file server for sideloading the add-in in
 * development. Uses office-addin-dev-certs so Word (and the OS) trust the
 * certificate, and serves the project root as static files - no bundler,
 * no build step. This is the only "server" involved: it just serves local
 * files over HTTPS; the add-in itself makes no network calls of its own.
 */
const fs = require("fs");
const path = require("path");
const https = require("https");
const { getHttpsServerOptions } = require("office-addin-dev-certs");

const PORT = process.env.PORT || 3000;
const ROOT = __dirname;

const MIME_TYPES = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".xml": "application/xml",
  ".ico": "image/x-icon",
};

function sendFile(res, filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const contentType = MIME_TYPES[ext] || "application/octet-stream";
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" });
      res.end("Not found: " + filePath);
      return;
    }
    res.writeHead(200, { "Content-Type": contentType });
    res.end(data);
  });
}

async function main() {
  const options = await getHttpsServerOptions();

  const server = https.createServer(options, (req, res) => {
    let urlPath = decodeURIComponent(req.url.split("?")[0]);
    if (urlPath === "/") urlPath = "/src/taskpane/taskpane.html";
    const filePath = path.join(ROOT, urlPath);

    // Prevent escaping the project root.
    if (!filePath.startsWith(ROOT)) {
      res.writeHead(403);
      res.end("Forbidden");
      return;
    }

    fs.stat(filePath, (err, stats) => {
      if (!err && stats.isFile()) {
        sendFile(res, filePath);
      } else {
        res.writeHead(404, { "Content-Type": "text/plain" });
        res.end("Not found");
      }
    });
  });

  server.listen(PORT, () => {
    console.log(`IDEA Lab Script Checker dev server running at https://localhost:${PORT}`);
    console.log(`Manifest: ${path.join(ROOT, "manifest.xml")}`);
  });
}

main().catch((err) => {
  console.error("Failed to start dev server:", err);
  console.error(
    "If this is the first time running the server, install the dev certificate with:\n" +
      "  npx office-addin-dev-certs install"
  );
  process.exit(1);
});
