// Serve this repo's manifests over HTTPS on 127.0.0.1:3001 with CORS "*", using
// the office-addin-dev-certs pair, so Office on the web can register the add-in
// through the wdaddindevserverport / wdaddinmanifestfile query parameters.
// Owns nothing else: two files, no writes, no state.
import { createServer } from "node:https";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { homedir } from "node:os";
import { fileURLToPath } from "node:url";

const REPO = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const CERTS = join(homedir(), ".office-addin-dev-certs");
const PORT = Number(process.env.RIG_MANIFEST_PORT ?? 3001);
const ALLOWED = new Set(["/manifest.prod.xml", "/manifest.xml"]);
const server = createServer(
  {
    key: readFileSync(join(CERTS, "localhost.key")),
    cert: readFileSync(join(CERTS, "localhost.crt")),
  },
  (req, res) => {
    const path = new URL(req.url, "https://x").pathname;
    const headers = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Headers": "*",
      "Cache-Control": "no-store",
    };
    if (req.method === "OPTIONS") {
      res.writeHead(204, headers);
      return res.end();
    }
    if (!ALLOWED.has(path)) {
      res.writeHead(404, headers);
      return res.end("not found");
    }
    res.writeHead(200, {
      ...headers,
      "Content-Type": "application/xml; charset=utf-8",
    });
    res.end(readFileSync(join(REPO, path)));
    console.log(
      new Date().toISOString(),
      req.method,
      path,
      req.headers.origin ?? "",
    );
  },
);
server.listen(PORT, "127.0.0.1", () =>
  console.log(`manifest server on https://127.0.0.1:${PORT} (serving ${REPO})`),
);
