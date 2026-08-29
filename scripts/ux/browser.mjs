// Chromium for the UX gate: finds playwright-core and a headless shell on
// this machine and launches one. Neither is a project dependency (the gate is
// dev tooling, not shipped code), so both fall back to the paths a plain
// `npx playwright install chromium-headless-shell` leaves behind. Owns
// nothing else: callers get a browser and close it themselves.

import { createRequire } from "node:module";
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import path from "node:path";

const DEFAULT_PLAYWRIGHT_DIR = path.join(
  homedir(),
  ".npm/_npx/31e32ef8478fbf80/node_modules",
);
const DEFAULT_CHROMIUM = path.join(
  homedir(),
  "Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell",
);

export function loadPlaywright() {
  // Prefer a real local install if one ever appears; fall back to the global
  // npx cache this environment already has Chromium's headless shell paired
  // with. createRequire is required because NODE_PATH is a CommonJS-only
  // resolution mechanism - ESM `import` never honours it.
  const require = createRequire(import.meta.url);
  try {
    return require("playwright-core");
  } catch {
    const dir = process.env.UX_CHECK_PLAYWRIGHT_DIR ?? DEFAULT_PLAYWRIGHT_DIR;
    const anchor = path.join(path.dirname(dir), "_ux-check-anchor.js");
    try {
      return createRequire(anchor)("playwright-core");
    } catch (error) {
      throw new Error(
        `playwright-core not found on NODE_PATH or at ${dir}. Install it ` +
          `or set UX_CHECK_PLAYWRIGHT_DIR to a node_modules directory that ` +
          `has it.\n${String(error)}`,
      );
    }
  }
}

export function resolveChromium() {
  const exe = process.env.UX_CHECK_CHROMIUM ?? DEFAULT_CHROMIUM;
  if (!existsSync(exe)) {
    throw new Error(
      `chrome-headless-shell not found at ${exe}. Run ` +
        `"npx playwright install chromium-headless-shell" or set ` +
        `UX_CHECK_CHROMIUM to its path.`,
    );
  }
  return exe;
}

export async function launchBrowser() {
  const { chromium } = loadPlaywright();
  return chromium.launch({ headless: true, executablePath: resolveChromium() });
}
