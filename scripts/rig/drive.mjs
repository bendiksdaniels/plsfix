// Drive the scratch Chrome (CDP on 127.0.0.1:9222) with one Playwright snippet
// per run, against Excel and PowerPoint for the web with the add-in registered.
// Usage: node scripts/rig/drive.mjs '<async (page, ctx, pages, shot, h) => {...}>' [pageIndex]
//        node scripts/rig/drive.mjs @scripts/rig/snippets/pane-open.js [pageIndex]
// Owns the CDP connection and the screenshot folder (.superpowers/rig, git-ignored).
import { mkdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { loadPlaywright } from "../ux/browser.mjs";
import * as h from "./helpers.mjs";

const CDP = process.env.RIG_CDP ?? "http://127.0.0.1:9222";
const [, , codeArg, pageIndexArg] = process.argv;
if (!codeArg) {
  console.error(
    "usage: node scripts/rig/drive.mjs '<snippet>' | @file [pageIndex]",
  );
  process.exit(2);
}
const raw = codeArg.startsWith("@")
  ? readFileSync(codeArg.slice(1), "utf8")
  : codeArg;
// A snippet file ends the way prettier leaves it, with a semicolon after the
// arrow function; wrapped in parentheses that is a syntax error, so it goes.
const code = raw.trim().replace(/;$/, "");
const { chromium } = loadPlaywright();
const browser = await chromium.connectOverCDP(CDP).catch((error) => {
  console.error(
    `no scratch Chrome at ${CDP} (${error.message.split("\n")[0]}); see scripts/rig/README.md`,
  );
  process.exit(2);
});
const ctx = browser.contexts()[0];
const pages = ctx.pages().filter((p) => !p.url().startsWith("chrome"));
const page = pages[Number(pageIndexArg ?? 0)] ?? pages[0];
await page.waitForTimeout(800); // let out-of-process iframes attach
const shot = async (name, opts = {}) => {
  mkdirSync(h.shotsDir(), { recursive: true });
  const path = join(h.shotsDir(), `${name}.png`);
  await page.screenshot({ path, ...opts });
  return path;
};
try {
  const fn = (0, eval)(`(${code})`);
  const result = await fn(page, ctx, pages, shot, h);
  console.log(
    typeof result === "string" ? result : JSON.stringify(result, null, 1),
  );
} catch (error) {
  console.error("ERR", error.message.split("\n").slice(0, 6).join("\n"));
  process.exitCode = 1;
} finally {
  await browser.close(); // disconnects from the scratch Chrome, never closes it
}
