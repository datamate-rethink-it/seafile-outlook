/**
 * Simple HTTPS dev server for testing the Seafile Outlook Add-in.
 * Serves the project files over https://localhost:3000.
 * Includes a reverse proxy for Seafile API calls to avoid CORS issues.
 *
 * Usage:
 *   1. Generate certificates: ./gen-certs.sh
 *   2. Start server: node dev/server.js
 *   3. Open https://localhost:3000 in browser and accept the self-signed cert
 *   4. Sideload manifest.xml in Outlook Web (https://aka.ms/olksideload)
 *
 * The proxy works by rewriting API calls:
 *   https://localhost:3000/seafile-proxy/https://cloud.seafile.com/api2/...
 *   → proxied to https://cloud.seafile.com/api2/...
 */

const https = require("https");
const http = require("http");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");

const PORT = 3000;
const ROOT = path.resolve(__dirname, "..");
const PROXY_PREFIX = "/seafile-proxy/";

// MIME types
const MIME = {
  ".html": "text/html",
  ".js": "application/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".md": "text/markdown",
  ".xml": "application/xml",
};

// Check for certificates
const certPath = path.join(__dirname, "cert.pem");
const keyPath = path.join(__dirname, "key.pem");

if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
  console.error("Certificates not found. Run ./dev/gen-certs.sh first.");
  process.exit(1);
}

/**
 * Proxy a request to the target Seafile server.
 * URL format: /seafile-proxy/{targetUrl}
 * e.g. /seafile-proxy/https://cloud.seafile.com/api2/auth-token/
 */
function handleProxy(req, res) {
  const targetUrl = req.url.substring(PROXY_PREFIX.length);
  if (!targetUrl || (!targetUrl.startsWith("http://") && !targetUrl.startsWith("https://"))) {
    res.writeHead(400);
    res.end("Bad proxy request: missing or invalid target URL");
    return;
  }

  let parsed;
  try {
    parsed = new URL(targetUrl);
  } catch {
    res.writeHead(400);
    res.end("Bad proxy request: invalid URL");
    return;
  }

  // Collect request body
  const bodyChunks = [];
  req.on("data", (chunk) => bodyChunks.push(chunk));
  req.on("end", () => {
    const body = Buffer.concat(bodyChunks);

    // Forward headers (except host and origin)
    const headers = {};
    for (const [key, val] of Object.entries(req.headers)) {
      const lower = key.toLowerCase();
      if (lower === "host" || lower === "origin" || lower === "referer") continue;
      headers[key] = val;
    }
    headers["host"] = parsed.host;

    const options = {
      hostname: parsed.hostname,
      port: parsed.port || (parsed.protocol === "https:" ? 443 : 80),
      path: parsed.pathname + parsed.search,
      method: req.method,
      headers,
    };

    const transport = parsed.protocol === "https:" ? https : http;
    const proxyReq = transport.request(options, (proxyRes) => {
      // Set CORS headers on the response
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.setHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
      res.setHeader("Access-Control-Allow-Headers", "Authorization, Content-Type, X-SEAFILE-OTP");

      res.writeHead(proxyRes.statusCode, proxyRes.headers);
      proxyRes.pipe(res);
    });

    proxyReq.on("error", (err) => {
      console.error("Proxy error:", err.message);
      res.writeHead(502);
      res.end(`Proxy error: ${err.message}`);
    });

    if (body.length > 0) {
      proxyReq.write(body);
    }
    proxyReq.end();
  });
}

const server = https.createServer(
  {
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  },
  (req, res) => {
    // Handle CORS preflight for proxy
    if (req.method === "OPTIONS") {
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.setHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
      res.setHeader("Access-Control-Allow-Headers", "Authorization, Content-Type, X-SEAFILE-OTP");
      res.writeHead(204);
      res.end();
      return;
    }

    // Proxy API requests
    if (req.url.startsWith(PROXY_PREFIX)) {
      handleProxy(req, res);
      return;
    }

    // Serve static files
    let url = req.url.split("?")[0];
    if (url === "/") url = "/compose/compose.html";

    const filePath = path.join(ROOT, url);

    if (!filePath.startsWith(ROOT)) {
      res.writeHead(403);
      res.end("Forbidden");
      return;
    }

    fs.readFile(filePath, (err, data) => {
      if (err) {
        res.writeHead(404);
        res.end(`Not found: ${url}`);
        return;
      }

      const ext = path.extname(filePath);
      const mime = MIME[ext] || "application/octet-stream";
      res.writeHead(200, { "Content-Type": mime });
      res.end(data);
    });
  }
);

server.listen(PORT, () => {
  console.log(`\n  Seafile Outlook Add-in dev server running at:`);
  console.log(`  https://localhost:${PORT}\n`);
  console.log(`  Task panes:`);
  console.log(`    Compose:  https://localhost:${PORT}/compose/compose.html`);
  console.log(`    Read:     https://localhost:${PORT}/read/read.html`);
  console.log(`    Settings: https://localhost:${PORT}/settings/settings.html`);
  console.log(`  Proxy:`);
  console.log(`    ${PROXY_PREFIX}{seafile-server-url}/api/...\n`);
});
