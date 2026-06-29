import { createServer } from "node:https";
import { existsSync, readFileSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL(".", import.meta.url));
const certDir = join(root, "work", "certs");
const keyPath = join(certDir, "localhost-key.pem");
const certPath = join(certDir, "localhost-cert.pem");
const port = Number(process.env.PORT || 3000);

await mkdir(certDir, { recursive: true });

if (!existsSync(keyPath) || !existsSync(certPath)) {
  execFileSync("openssl", [
    "req",
    "-x509",
    "-newkey",
    "rsa:2048",
    "-nodes",
    "-sha256",
    "-days",
    "365",
    "-subj",
    "/CN=localhost",
    "-keyout",
    keyPath,
    "-out",
    certPath
  ]);
}

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml; charset=utf-8",
  ".xml": "application/xml; charset=utf-8"
};

function resolveRequestPath(url) {
  const requested = new URL(url, `https://localhost:${port}`).pathname;
  const clean = normalize(decodeURIComponent(requested)).replace(/^(\.\.(\/|\\|$))+/, "");
  const path = clean === "/" ? "/src/taskpane.html" : clean;
  return join(root, path);
}

createServer(
  {
    key: readFileSync(keyPath),
    cert: readFileSync(certPath)
  },
  (request, response) => {
    const filePath = resolveRequestPath(request.url || "/");

    try {
      const body = readFileSync(filePath);
      response.writeHead(200, {
        "Content-Type": mimeTypes[extname(filePath)] || "application/octet-stream",
        "Cache-Control": "no-store"
      });
      response.end(body);
    } catch {
      response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("Not found");
    }
  }
).listen(port, () => {
  console.log(`PowerPoint add-in server: https://localhost:${port}/src/taskpane.html`);
  console.log(`Manifest: ${join(root, "manifest.xml")}`);
});
