import http from "node:http";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT || 5173);
const ROOT = __dirname;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".csv": "text/csv; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml; charset=utf-8",
};

function safeJoin(rootDir, reqPath) {
  const decoded = decodeURIComponent(reqPath.split("?")[0]);
  const clean = decoded.replace(/\\/g, "/");
  const joined = path.join(rootDir, clean);
  const resolved = path.resolve(joined);
  const resolvedRoot = path.resolve(rootDir);
  if (!resolved.startsWith(resolvedRoot)) return null;
  return resolved;
}

const server = http.createServer(async (req, res) => {
  try {
    const urlPath = req.url === "/" ? "/index.html" : req.url || "/index.html";
    const filePath = safeJoin(ROOT, urlPath);
    if (!filePath) {
      res.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Bad request");
      return;
    }

    let st;
    try {
      st = await stat(filePath);
    } catch {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not found");
      return;
    }

    if (st.isDirectory()) {
      res.writeHead(302, { Location: "/index.html" });
      res.end();
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME[ext] || "application/octet-stream";
    const body = await readFile(filePath);
    res.writeHead(200, { "Content-Type": contentType, "Cache-Control": "no-store" });
    res.end(body);
  } catch (e) {
    res.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    res.end(e?.message ?? "Internal error");
  }
});

function listenWithFallback(startPort) {
  let port = startPort;
  const tryListen = () => {
    server.once("error", (err) => {
      if (err && err.code === "EADDRINUSE") {
        port += 1;
        tryListen();
        return;
      }
      throw err;
    });
    server.listen(port, "127.0.0.1", () => {
      // eslint-disable-next-line no-console
      console.log(`Dashboard: http://127.0.0.1:${port}/`);
    });
  };
  tryListen();
}

listenWithFallback(PORT);
