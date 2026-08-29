#!/usr/bin/env node
// UX pass over both task panes at narrow widths (backlog C3). Boots the two
// panes (taskpane.html / pptpane.html) against a running `npm run dev`
// server, walks every tab at three viewport widths (320/360/420 x 720),
// seeds the Links/Inbox tables with sample data through the real render
// functions (src/pane/links-list.ts, src/ppt/views.ts) via a dynamic
// `import()` served live by Vite, screenshots each combination and checks
// it for five defect classes:
//   1. an element's box overflowing the viewport horizontally
//   2. a button/input/select shorter than a 32px tap target
//      (native checkboxes, color pickers and hidden file inputs are exempt -
//      they follow their own, smaller platform convention)
//   3. a button's own label ellipsis-clipped (scrollWidth > clientWidth on
//      an overflow:hidden/text-overflow:ellipsis element that is a <button>
//      or sits inside one) - deliberately truncated *data* in a non-button
//      cell (a workbook name, a cell address) is not this bug
//   4. text rendered below an 11px computed font size
//   5. real Tab-key focus order through the first 12 stops, flagging any
//      stop that is not actually visible, plus document.body.scrollWidth
//      exceeding the viewport (page-level horizontal scroll)
//
// Usage:
//   npm run dev -- --port 3110          # in one terminal
//   node scripts/ux-check.mjs --port 3110
//
// Requires a Chromium build and playwright-core on NODE_PATH; neither is a
// project dependency (this is a throwaway dev script, not shipped code).
// Both default to the paths a plain `npx playwright install chromium-headless-shell`
// leaves on this machine; override with UX_CHECK_PLAYWRIGHT_DIR /
// UX_CHECK_CHROMIUM if that ever moves. Exits 0 with an empty defect list,
// 1 otherwise - safe to use as a gate.

import { createRequire } from "node:module";
import { existsSync, mkdirSync } from "node:fs";
import { homedir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = fileURLToPath(new URL("..", import.meta.url));
const OUT_DIR = path.join(ROOT, ".superpowers/sdd/2026-08-28-ppt-links/ux");

const DEFAULT_PLAYWRIGHT_DIR = path.join(
  homedir(),
  ".npm/_npx/31e32ef8478fbf80/node_modules",
);
const DEFAULT_CHROMIUM = path.join(
  homedir(),
  "Library/Caches/ms-playwright/chromium_headless_shell-1234/chrome-headless-shell-mac-arm64/chrome-headless-shell",
);

const WIDTHS = [320, 360, 420];
const HEIGHT = 720;
const TAB_PRESSES = 12;
const MIN_TAP_HEIGHT = 32;
const MIN_FONT_PX = 11;

function loadPlaywright() {
  // Prefer a real local install if one ever appears; fall back to the global
  // npx cache this environment already has Chromium's headless shell paired
  // with (see header). createRequire is required because NODE_PATH is a
  // CommonJS-only resolution mechanism - ESM `import` never honours it.
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

function resolveChromium() {
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

function parsePort(argv) {
  const flagIndex = argv.indexOf("--port");
  if (flagIndex === -1) return 3000;
  const value = argv[flagIndex + 1];
  const port = Number(value);
  if (!Number.isInteger(port) || port <= 0) {
    throw new Error(`--port needs an integer, got ${String(value)}`);
  }
  return port;
}

const PANES = [
  {
    file: "taskpane.html",
    slug: "taskpane",
    seed: async (page) => {
      await page.evaluate(async () => {
        const mod = await import("/src/pane/links-list.ts");
        const tbody = document.getElementById("workbook-links");
        if (!tbody) throw new Error("#workbook-links missing");
        const rows = [
          {
            entry: {
              id: "a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1",
              kind: "range",
              anchor: "PLSFIX_LINK_a1a1a1a1",
              label: "Revenue bridge FY25-FY26, EUR '000",
              token: "tok-1",
              createdAt: "2026-08-01T09:00:00Z",
              lastPushedAt: "2026-08-27T14:32:00Z",
              rev: 3,
            },
            source: "ok",
          },
          {
            entry: {
              id: "b2b2b2b2b2b2b2b2b2b2b2b2b2b2b2b2",
              kind: "chart",
              anchor: "PLSFIX_LINK_b2b2b2b2",
              label:
                "EBITDA margin trend, quarterly, all segments combined and restated",
              token: "tok-2",
              createdAt: "2026-07-15T09:00:00Z",
              lastPushedAt: null,
              rev: 1,
            },
            source: "ok",
          },
          {
            entry: {
              id: "c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3",
              kind: "range",
              anchor: "PLSFIX_LINK_c3c3c3c3",
              label: "Net debt bridge",
              token: "tok-3",
              createdAt: "2026-06-01T09:00:00Z",
              lastPushedAt: "2026-06-02T09:00:00Z",
              rev: 1,
            },
            source: "missing",
          },
          {
            entry: {
              id: "d4d4d4d4d4d4d4d4d4d4d4d4d4d4d4d4",
              kind: "chart",
              anchor: "PLSFIX_LINK_d4d4d4d4",
              label: "Segment revenue mix",
              token: "tok-4",
              createdAt: "2026-08-20T09:00:00Z",
              lastPushedAt: "2026-08-28T08:00:00Z",
              rev: 2,
            },
            source: "ok",
          },
        ];
        const selected = new Set([rows[0].entry.id]);
        mod.renderWorkbookLinks(tbody, rows, selected, () => undefined);
      });
    },
  },
  {
    file: "pptpane.html",
    slug: "pptpane",
    seed: async (page) => {
      await page.evaluate(async () => {
        const mod = await import("/src/ppt/views.ts");
        const now = Date.now() / 1000;
        const linkRows = document.getElementById("link-rows");
        if (!linkRows) throw new Error("#link-rows missing");
        const rows = [
          {
            key: "k1",
            slide: 3,
            label: "Revenue bridge FY25-FY26, EUR '000, full detail",
            source: "Q4 2026 Model - Consolidated Group View.xlsx",
            status: "current",
            pushedAt: now - 3600,
            selected: true,
          },
          {
            key: "k2",
            slide: 12,
            label: "EBITDA margin trend",
            source: "Model.xlsx",
            status: "updateAvailable",
            pushedAt: now - 90000,
            selected: false,
          },
          {
            key: "k3",
            slide: 5,
            label: "Segment revenue mix chart",
            source: "VeryLongWorkbookNameThatMightOverflowTheColumn.xlsx",
            status: "missing",
            pushedAt: null,
            selected: false,
          },
          {
            key: "k4",
            slide: 21,
            label: "Net debt bridge",
            source: "Model.xlsx",
            status: "wrongKey",
            pushedAt: now - 200000,
            selected: false,
          },
        ];
        mod.renderLinkRows(linkRows, rows, () => undefined);
        const linksEmpty = document.getElementById("links-empty");
        if (linksEmpty) linksEmpty.hidden = rows.length > 0;

        const inboxList = document.getElementById("inbox-list");
        if (!inboxList) throw new Error("#inbox-list missing");
        const items = [
          {
            id: "i1",
            token: "tk1",
            kind: "range",
            label: "Revenue bridge FY25-FY26, EUR '000",
            src: {
              workbook: "Q4 2026 Model - Consolidated Group View.xlsx",
              sheet: "Summary",
              ref: "B2:H14",
              anchor: "PLSFIX_LINK_i1",
            },
            createdAt: new Date(Date.now() - 3600_000).toISOString(),
          },
          {
            id: "i2",
            token: "tk2",
            kind: "chart",
            label: "EBITDA trend",
            src: {
              workbook: "Model.xlsx",
              sheet: "Charts",
              ref: "Chart 1",
              anchor: "PLSFIX_LINK_i2",
            },
            createdAt: new Date(Date.now() - 86_400_000 * 2).toISOString(),
          },
        ];
        mod.renderInbox(inboxList, items, () => undefined);
        inboxList.hidden = false;
        const unpaired = document.getElementById("inbox-unpaired");
        if (unpaired) unpaired.hidden = true;
      });
    },
  },
];

// Runs inside the page. Must be fully self-contained (no closures over the
// outer Node scope survive serialization).
function evaluateDefects() {
  const vw = window.innerWidth;
  const EPS_ = 0.5;
  const MIN_TAP_HEIGHT_ = 32;
  const MIN_FONT_PX_ = 11;

  function describe(el) {
    if (el.id) return `#${el.id}`;
    let s = el.tagName.toLowerCase();
    if (typeof el.className === "string" && el.className.trim()) {
      s += "." + el.className.trim().split(/\s+/).join(".");
    }
    const parent = el.parentElement;
    if (parent) {
      const idx = Array.prototype.indexOf.call(parent.children, el);
      s += `[${String(idx)}]`;
    }
    return s;
  }

  function isRendered(el) {
    const rects = el.getClientRects();
    if (rects.length === 0) return false;
    const r = rects[0];
    if (r.width === 0 && r.height === 0) return false;
    const cs = getComputedStyle(el);
    if (cs.visibility === "hidden" || cs.display === "none") return false;
    return true;
  }

  const overflow = [];
  const shortControls = [];
  const clippedButtons = [];
  const smallText = [];

  for (const el of document.querySelectorAll("body *")) {
    if (!isRendered(el)) continue;
    const rect = el.getBoundingClientRect();
    const tag = el.tagName;
    const type = (el.getAttribute("type") || "").toLowerCase();

    if (rect.right > vw + EPS_) {
      overflow.push(
        `${describe(el)} right=${rect.right.toFixed(1)} viewport=${String(vw)}`,
      );
    }

    const isTapControl =
      tag === "BUTTON" ||
      tag === "SELECT" ||
      (tag === "INPUT" &&
        !["checkbox", "color", "file", "hidden"].includes(type));
    if (
      isTapControl &&
      rect.height > 0 &&
      rect.height < MIN_TAP_HEIGHT_ - EPS_
    ) {
      shortControls.push(`${describe(el)} height=${rect.height.toFixed(1)}`);
    }

    if (el.scrollWidth > el.clientWidth + EPS_) {
      const cs = getComputedStyle(el);
      const clips = cs.overflowX === "hidden" || cs.textOverflow === "ellipsis";
      const isButtonish = tag === "BUTTON" || el.closest("button") !== null;
      if (clips && isButtonish) {
        const text = (el.textContent || "").trim().slice(0, 60);
        clippedButtons.push(
          `${describe(el)} scrollWidth=${String(el.scrollWidth)} clientWidth=${String(el.clientWidth)} text="${text}"`,
        );
      }
    }

    let hasOwnText = false;
    for (const node of el.childNodes) {
      if (node.nodeType === 3 && (node.textContent || "").trim().length > 0) {
        hasOwnText = true;
        break;
      }
    }
    if (hasOwnText) {
      // No EPS grace band here: font-size is an authored exact value (this
      // codebase has a real 10.5px bucket), not a layout-rounding artefact
      // like a bounding-box edge - a tiny float-representation guard is
      // still worth keeping so 11.0000000001px from a DPI transform isn't
      // flagged.
      const fontPx = parseFloat(getComputedStyle(el).fontSize);
      if (fontPx < MIN_FONT_PX_ - 0.05) {
        const text = (el.textContent || "").trim().slice(0, 40);
        smallText.push(`${describe(el)} font=${fontPx}px text="${text}"`);
      }
    }
  }

  return {
    overflow,
    shortControls,
    clippedButtons,
    smallText,
    bodyScrollWidth: document.body.scrollWidth,
    bodyScrolls: document.body.scrollWidth > vw + EPS_,
    viewportWidth: vw,
  };
}

async function checkFocusOrder(page) {
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
      const visible =
        rects.length > 0 &&
        !(rects[0].width === 0 && rects[0].height === 0) &&
        getComputedStyle(el).visibility !== "hidden" &&
        getComputedStyle(el).display !== "none";
      const desc = el.id
        ? `#${el.id}`
        : el.tagName.toLowerCase() +
          (typeof el.className === "string" && el.className.trim()
            ? "." + el.className.trim().split(/\s+/).join(".")
            : "");
      return { desc, visible };
    });
    if (stop === null) break;
    stops.push(stop);
  }
  return stops;
}

function discoverTabs(page) {
  return page.evaluate(() => {
    return Array.from(document.querySelectorAll('[role="tab"]')).map((tab) => ({
      id: tab.id,
      slug: tab.id.replace(/^tab-/, ""),
    }));
  });
}

async function run() {
  const argv = process.argv.slice(2);
  const port = parsePort(argv);
  const baseUrl = `https://localhost:${String(port)}`;

  const { chromium } = loadPlaywright();
  const executablePath = resolveChromium();
  mkdirSync(OUT_DIR, { recursive: true });

  const browser = await chromium.launch({ headless: true, executablePath });
  const combos = [];

  try {
    for (const pane of PANES) {
      const context = await browser.newContext({ ignoreHTTPSErrors: true });
      const page = await context.newPage();
      page.on("pageerror", (error) => {
        console.error(`  [page error, ${pane.file}] ${String(error)}`);
      });

      await page.goto(`${baseUrl}/${pane.file}`, { waitUntil: "networkidle" });
      await page.waitForFunction(
        () =>
          (document.getElementById("app-version")?.textContent ?? "") !== "",
        { timeout: 15000 },
      );
      await pane.seed(page);

      const tabs = await discoverTabs(page);
      for (const width of WIDTHS) {
        await page.setViewportSize({ width, height: HEIGHT });
        for (const tab of tabs) {
          await page.click(`#${tab.id}`);
          await page.waitForTimeout(60);

          const screenshotPath = path.join(
            OUT_DIR,
            `${pane.slug}-${tab.slug}-${String(width)}.png`,
          );
          await page.screenshot({ path: screenshotPath });

          const defects = await page.evaluate(evaluateDefects);
          const focusOrder = await checkFocusOrder(page);
          const invisibleFocus = focusOrder.filter((s) => !s.visible);

          combos.push({
            pane: pane.slug,
            tab: tab.slug,
            width,
            screenshot: path.relative(ROOT, screenshotPath),
            defects,
            focusOrder,
            invisibleFocus,
          });
        }
      }
      await context.close();
    }
  } finally {
    await browser.close();
  }

  let totalDefects = 0;
  for (const combo of combos) {
    const lines = [];
    for (const item of combo.defects.overflow) {
      lines.push(`  overflow: ${item}`);
    }
    for (const item of combo.defects.shortControls) {
      lines.push(`  short-control (<${String(MIN_TAP_HEIGHT)}px): ${item}`);
    }
    for (const item of combo.defects.clippedButtons) {
      lines.push(`  clipped-button-text: ${item}`);
    }
    for (const item of combo.defects.smallText) {
      lines.push(`  small-text (<${String(MIN_FONT_PX)}px): ${item}`);
    }
    if (combo.defects.bodyScrolls) {
      lines.push(
        `  body-scroll: scrollWidth=${String(combo.defects.bodyScrollWidth)} > viewport=${String(combo.defects.viewportWidth)}`,
      );
    }
    for (const stop of combo.invisibleFocus) {
      lines.push(`  invisible-focus: ${stop.desc}`);
    }

    const header = `=== ${combo.pane} / ${combo.tab} / ${String(combo.width)}px ===`;
    if (lines.length === 0) {
      console.log(`${header} clean (screenshot: ${combo.screenshot})`);
    } else {
      console.log(
        `${header} ${String(lines.length)} defect(s) (screenshot: ${combo.screenshot})`,
      );
      for (const line of lines) console.log(line);
      totalDefects += lines.length;
    }
  }

  console.log("");
  console.log(
    `focus order sample (${PANES[0]?.slug ?? "?"} / first tab / ${String(WIDTHS[0])}px):`,
  );
  const sample = combos[0];
  if (sample) {
    console.log(
      "  " +
        sample.focusOrder
          .map((s) => s.desc + (s.visible ? "" : " [INVISIBLE]"))
          .join(" -> "),
    );
  }

  console.log("");
  console.log(
    totalDefects === 0
      ? `PASS - 0 defects across ${String(combos.length)} pane/tab/width combinations`
      : `FAIL - ${String(totalDefects)} defect(s) across ${String(combos.length)} combinations`,
  );

  process.exitCode = totalDefects === 0 ? 0 : 1;
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
