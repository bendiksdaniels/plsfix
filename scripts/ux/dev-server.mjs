// The pane host for the UX gate. Reuses a dev server already listening on the
// port, otherwise starts its own vite on a free one and stops it again, so
// `npm run ux:check` is a single command with no terminal to set up first.
// Owns the child process and nothing else; the https certificate is whatever
// vite.config.ts hands out.

import { spawn } from "node:child_process";
import { connect } from "node:net";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = fileURLToPath(new URL("../..", import.meta.url));
const BOOT_TIMEOUT_MS = 60_000;
const POLL_MS = 250;

function isListening(port) {
  return new Promise((resolve) => {
    const socket = connect({ port, host: "127.0.0.1" });
    const settle = (value) => {
      socket.destroy();
      resolve(value);
    };
    socket.once("connect", () => settle(true));
    socket.once("error", () => settle(false));
    socket.setTimeout(1000, () => settle(false));
  });
}

async function waitForPort(port, child) {
  const deadline = Date.now() + BOOT_TIMEOUT_MS;
  while (Date.now() < deadline) {
    if (await isListening(port)) return;
    if (child.exitCode !== null) {
      throw new Error(`vite exited with code ${String(child.exitCode)}`);
    }
    await new Promise((resolve) => setTimeout(resolve, POLL_MS));
  }
  throw new Error(`vite did not listen on ${String(port)} in time`);
}

// A caller-named port is a promise that something is already there; an unnamed
// one means we own the server, on a port unlikely to collide with a sideload.
export async function ensureServer(port) {
  if (await isListening(port)) {
    return { url: `https://localhost:${String(port)}`, stop: () => undefined };
  }
  const vite = path.join(ROOT, "node_modules/.bin/vite");
  const child = spawn(vite, ["--port", String(port), "--strictPort"], {
    cwd: ROOT,
    stdio: "ignore",
  });
  try {
    await waitForPort(port, child);
  } catch (error) {
    child.kill("SIGTERM");
    throw error;
  }
  return {
    url: `https://localhost:${String(port)}`,
    stop: () => child.kill("SIGTERM"),
  };
}
