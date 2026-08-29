#!/usr/bin/env node
// UX gate over both task panes at the widths Office actually gives a pane
// (320/360/420/500 CSS px). Boots a vite dev server if one is not already
// listening, walks every tab of every pane in every seeded state, screenshots
// each combination and fails on any of: an element past the viewport edge, a
// body/section/tab-bar that scrolls horizontally, a tab pushed out of the
// strip or its label clipped, a control under a 32px tap target, a clipped
// button label, text below 11px, a toast covering the tab strip, or a Tab stop
// that is not visible.
//
// Usage:
//   npm run ux:check                      # boots its own server
//   node scripts/ux-check.mjs --port 3000 # reuse a running `npm run dev`
//   node scripts/ux-check.mjs --out docs/ux/2026-08-29/after --states filled
//
// Requires playwright-core and a Chromium headless shell on this machine; see
// scripts/ux/browser.mjs. Exits 0 with an empty defect list, 1 otherwise.

import { mkdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { launchBrowser } from "./ux/browser.mjs";
import { ensureServer } from "./ux/dev-server.mjs";
import { PANES } from "./ux/panes.mjs";
import { printCombos, printSummary } from "./ux/report.mjs";

const ROOT = fileURLToPath(new URL("..", import.meta.url));
const PROBE = path.join(ROOT, "scripts/ux/page-probe.js");
const WIDTHS = [320, 360, 420, 500];
const HEIGHT = 720;
const TAB_PRESSES = 12;
const SETTLE_MS = 260;
const OWN_PORT = 3131;

function flag(argv, name, fallback) {
  const at = argv.indexOf(`--${name}`);
  return at === -1 ? fallback : argv[at + 1];
}

function parseArgs(argv) {
  const port = Number(flag(argv, "port", OWN_PORT));
  if (!Number.isInteger(port) || port <= 0) {
    throw new Error(`--port needs an integer, got ${String(port)}`);
  }
  const only = flag(argv, "states", "");
  return {
    port,
    outDir: path.resolve(ROOT, flag(argv, "out", ".superpowers/ux")),
    states: only ? only.split(",") : null,
    widths: (flag(argv, "widths", "") || WIDTHS.join(","))
      .split(",")
      .map(Number),
  };
}

async function tabStops(page) {
  await page.evaluate(() => {
    if (document.activeElement instanceof HTMLElement) {
      document.activeElement.blur();
    }
  });
  const stops = [];
  for (let i = 0; i < TAB_PRESSES; i += 1) {
    await page.keyboard.press("Tab");
    const stop = await page.evaluate(() => {
      const el = document.activeElement;
      if (!el || el === document.body) return null;
      const rects = el.getClientRects();
      const cs = getComputedStyle(el);
      const visible =
        rects.length > 0 &&
        !(rects[0].width === 0 && rects[0].height === 0) &&
        cs.visibility !== "hidden" &&
        cs.display !== "none";
      const cls =
        typeof el.className === "string" && el.className.trim()
          ? "." + el.className.trim().split(/\s+/).join(".")
          : "";
      return {
        desc: el.id ? `#${el.id}` : el.tagName.toLowerCase() + cls,
        visible,
      };
    });
    if (stop === null) break;
    stops.push(stop);
  }
  return stops;
}

async function openPane(browser, baseUrl, pane) {
  const context = await browser.newContext({ ignoreHTTPSErrors: true });
  const page = await context.newPage();
  page.on("pageerror", (error) => {
    process.stderr.write(`  [page error, ${pane.file}] ${String(error)}\n`);
  });
  await page.goto(`${baseUrl}/${pane.file}`, { waitUntil: "networkidle" });
  await page.waitForFunction(
    () => (document.getElementById("app-version")?.textContent ?? "") !== "",
    { timeout: 15_000 },
  );
  await page.addScriptTag({ path: PROBE });
  return { context, page };
}

function tabsFor(page, state) {
  return page
    .evaluate(() =>
      Array.from(document.querySelectorAll("[role=tab]")).map((tab) => ({
        id: tab.id,
        slug: tab.id.replace(/^tab-/, ""),
      })),
    )
    .then((all) =>
      state.tabs ? all.filter((tab) => state.tabs.includes(tab.slug)) : all,
    );
}

async function measure(page, pane, state, tab, width, outDir) {
  await page.click(`#${tab.id}`);
  // The Tab-key walk of the previous combination scrolls; every shot is of the
  // top of the pane, and the settle covers the .15s state transitions.
  await page.evaluate(() => {
    window.scrollTo(0, 0);
  });
  await page.waitForTimeout(SETTLE_MS);
  const name = `${pane.slug}-${state.slug}-${tab.slug}-${String(width)}.png`;
  const screenshot = path.join(outDir, name);
  await page.screenshot({ path: screenshot });
  const defects = await page.evaluate(() => window.__ux.collect());
  const focusOrder = state.focus ? await tabStops(page) : [];
  return {
    pane: pane.slug,
    state: state.slug,
    tab: tab.slug,
    width,
    screenshot: path.relative(ROOT, screenshot),
    defects,
    focusOrder,
    invisibleFocus: focusOrder.filter((stop) => !stop.visible),
  };
}

async function walkPane(browser, baseUrl, pane, options) {
  const { context, page } = await openPane(browser, baseUrl, pane);
  const combos = [];
  try {
    for (const state of pane.states) {
      if (options.states && !options.states.includes(state.slug)) continue;
      await state.apply(page);
      const tabs = await tabsFor(page, state);
      for (const width of options.widths) {
        await page.setViewportSize({ width, height: HEIGHT });
        for (const tab of tabs) {
          combos.push(
            await measure(page, pane, state, tab, width, options.outDir),
          );
        }
      }
    }
  } finally {
    await context.close();
  }
  return combos;
}

async function run() {
  const options = parseArgs(process.argv.slice(2));
  mkdirSync(options.outDir, { recursive: true });
  const server = await ensureServer(options.port);
  const browser = await launchBrowser();
  const combos = [];
  try {
    for (const pane of PANES) {
      combos.push(...(await walkPane(browser, server.url, pane, options)));
    }
  } finally {
    await browser.close();
    server.stop();
  }
  const log = (line) => process.stdout.write(`${line}\n`);
  const total = printCombos(combos, log);
  printSummary(combos, total, log);
  process.exitCode = total === 0 ? 0 : 1;
}

run().catch((error) => {
  process.stderr.write(`${String(error)}\n`);
  process.exitCode = 1;
});
