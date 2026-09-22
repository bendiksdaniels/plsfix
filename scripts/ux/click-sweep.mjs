#!/usr/bin/env node
// Functional click-through gate for both panes, with NO Office host at all
// (the real office.js from the CDN, loaded outside any Office client:
// Office.onReady() resolves quickly with a host that is neither Excel nor
// PowerPoint, so the pane boots degraded - src/host-ready.ts, src/main.ts,
// src/ppt/main.ts). Every [data-action] button, tab, help "?" toggle,
// first-run card button, select and checkbox is driven with a real DOM
// click/change; every tool's own label is typed into the search box and its
// top result clicked.
//
// A "reaction" (scripts/ux/sweep-probe.js) is any of: document.body.innerHTML
// changed, an input/select/textarea's value or checked state changed, the
// active element changed, or document.title changed - between them these
// catch a toast, a re-rendered list, a hidden/aria/class flip and a focus
// move without enumerating every attribute a given fix might touch.
//
// FAILs on: an uncaught error or unhandled rejection (window.onerror/
// unhandledrejection, hooked the way src/ui/report.ts listens for them, plus
// Playwright's own pageerror/console as a second, injection-independent
// layer); no reaction within CLICK_BUDGET_MS; a tab leaving zero or two+
// panels visible; two help cards open at once; a search row landing on a
// different tab than its own label promised; two buttons sharing one
// data-action or first-run id (the mechanism behind a wrong-tool search
// click); a focus walk that never reaches a sane fraction of the page's own
// focusable controls (a trap).
//
// Usage: npm run ux:sweep [-- --port 3000] [-- --pane taskpane]

import path from "node:path";
import { fileURLToPath } from "node:url";

import { launchBrowser } from "./browser.mjs";
import { ensureServer } from "./dev-server.mjs";

const ROOT = fileURLToPath(new URL("../..", import.meta.url));
const PROBE = path.join(ROOT, "scripts/ux/sweep-probe.js");
const OWN_PORT = 3132;
const CLICK_BUDGET_MS = 500;
const BOOT_TIMEOUT_MS = 15_000;

const PANES = [
  { file: "taskpane.html", slug: "taskpane" },
  { file: "pptpane.html", slug: "pptpane" },
];

function flag(argv, name, fallback) {
  const at = argv.indexOf(`--${name}`);
  return at === -1 ? fallback : argv[at + 1];
}

function parseArgs(argv) {
  const port = Number(flag(argv, "port", OWN_PORT));
  if (!Number.isInteger(port) || port <= 0) {
    throw new Error(`--port needs an integer, got ${String(port)}`);
  }
  const onlyPane = flag(argv, "pane", null);
  const slugs = PANES.map((p) => p.slug);
  if (onlyPane !== null && !slugs.includes(onlyPane)) {
    throw new Error(
      `--pane "${onlyPane}" is not one of this sweep's panes. Valid: ${slugs.join(", ")}`,
    );
  }
  return {
    port,
    panes: onlyPane ? PANES.filter((p) => p.slug === onlyPane) : PANES,
  };
}

async function openPane(browser, baseUrl, pane) {
  const context = await browser.newContext({ ignoreHTTPSErrors: true });
  const page = await context.newPage();
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(String(error)));
  page.on("console", (message) => {
    if (message.type() === "error") pageErrors.push(message.text());
  });
  await page.goto(`${baseUrl}/${pane.file}`, { waitUntil: "domcontentloaded" });
  await page.addScriptTag({ path: PROBE });
  await page.evaluate(() => window.__sweep.installHooks());
  await page.waitForFunction(
    () => {
      const text = document.getElementById("connection-status")?.textContent;
      return Boolean(text) && text !== "Connecting";
    },
    { timeout: BOOT_TIMEOUT_MS },
  );
  return { context, page, pageErrors };
}

function clickKind(target) {
  if (target.kind === "checkbox" || target.kind === "select")
    return target.kind;
  return "click";
}

function label(paneSlug, target) {
  return `${paneSlug} | ${target.kind} ${target.selector} ("${target.label}")`;
}

// One target's click/change plus its own structural follow-up check (tabs:
// exactly one panel and it is the right one; help: never two cards open).
async function runTarget(page, paneSlug, target, failures) {
  const result = await page.evaluate(
    ([selector, kind, budget]) =>
      window.__sweep.clickAndObserve(selector, kind, budget),
    [target.selector, clickKind(target), CLICK_BUDGET_MS],
  );
  if (!result.found) {
    failures.push(`${label(paneSlug, target)} | control not found`);
    return;
  }
  if (result.skippedDisabled || result.skippedNoOptions) return;
  if (result.erred) {
    const errs = await page.evaluate(() => window.__sweep.errors());
    failures.push(
      `${label(paneSlug, target)} | ${errs[errs.length - 1] ?? "uncaught error"}`,
    );
    return;
  }
  if (!result.reacted) {
    failures.push(
      `${label(paneSlug, target)} | no reaction within ${String(CLICK_BUDGET_MS)}ms`,
    );
    return;
  }
  if (target.kind === "tab")
    await checkTabExclusivity(page, paneSlug, target, failures);
  if (target.kind === "help")
    await checkHelpExclusivity(page, paneSlug, target, failures);
}

async function checkTabExclusivity(page, paneSlug, target, failures) {
  const state = await page.evaluate(() => window.__sweep.tabState());
  const wantId = target.selector.slice(1);
  if (state.visiblePanels.length !== 1) {
    failures.push(
      `${label(paneSlug, target)} | ${String(state.visiblePanels.length)} panels visible, want 1`,
    );
  } else if (state.activeTabId !== wantId) {
    failures.push(
      `${label(paneSlug, target)} | active tab is ${String(state.activeTabId)}, want ${wantId}`,
    );
  }
}

async function checkHelpExclusivity(page, paneSlug, target, failures) {
  const open = await page.evaluate(() => window.__sweep.openHelpCards());
  if (open.length > 1) {
    failures.push(
      `${label(paneSlug, target)} | ${String(open.length)} help cards open at once: ${open.join(", ")}`,
    );
  }
}

// data-action/first-run-id collisions: the exact mechanism that would let a
// search row's click land on a different tool than the label it showed.
async function checkDuplicateActions(page, paneSlug, failures) {
  const dupes = await page.evaluate(() => window.__sweep.duplicateActionKeys());
  for (const key of dupes) {
    failures.push(
      `${paneSlug} | duplicate action key "${key}" shared by two buttons`,
    );
  }
}

async function checkSearchRows(page, paneSlug, targets, failures) {
  const hasSearch = await page.evaluate(() =>
    Boolean(document.getElementById("tool-search")),
  );
  if (!hasSearch) return;
  for (const target of targets.filter((t) => t.kind === "button" && t.label)) {
    const result = await page.evaluate(
      (text) => window.__sweep.checkSearchRow(text),
      target.label,
    );
    if (result.skipped || !result.found) continue;
    if (!result.ok) {
      failures.push(
        `${label(paneSlug, target)} | search row landed on "${result.landedTab}", label said "${result.rowTab}"`,
      );
    }
  }
}

const FOCUSABLE_SELECTOR =
  'button:not([disabled]), [href], input:not([disabled]):not([type="hidden"]), ' +
  'select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';

// Tabs through the whole page from a blurred state, identifying each stop by
// its index in the same focusable-elements list (never by id or tag name
// alone - most of these buttons share no id, and two anonymous buttons would
// otherwise collide into one "seen" entry and undercount). A trap shows up
// as visiting far fewer distinct controls than the page actually has; a few
// consecutive misses (focus landing on document.body) are tolerated before
// concluding focus has genuinely left the document.
async function checkFocusWalk(page, paneSlug, failures) {
  await page.evaluate(() => {
    if (document.activeElement instanceof HTMLElement)
      document.activeElement.blur();
  });
  const focusable = await page.evaluate(
    (selector) =>
      Array.from(document.querySelectorAll(selector)).filter(
        (el) => el.getClientRects().length > 0,
      ).length,
    FOCUSABLE_SELECTOR,
  );
  if (focusable === 0) return;
  const seen = new Set();
  const ceiling = focusable * 2 + 10;
  let misses = 0;
  for (let i = 0; i < ceiling && misses <= 3; i += 1) {
    await page.keyboard.press("Tab");
    const index = await page.evaluate(
      (selector) =>
        Array.from(document.querySelectorAll(selector)).indexOf(
          document.activeElement,
        ),
      FOCUSABLE_SELECTOR,
    );
    if (index === -1) {
      misses += 1;
      continue;
    }
    misses = 0;
    seen.add(index);
  }
  if (seen.size < focusable * 0.5) {
    failures.push(
      `${paneSlug} | focus walk reached ${String(seen.size)} of ~${String(focusable)} focusable controls - possible trap`,
    );
  }
}

async function walkPane(browser, baseUrl, pane) {
  const { context, page, pageErrors } = await openPane(browser, baseUrl, pane);
  const failures = [];
  let count = 0;
  try {
    const targets = await page.evaluate(() => window.__sweep.listTargets());
    for (const target of targets) {
      await runTarget(page, pane.slug, target, failures);
      count += 1;
    }
    await checkDuplicateActions(page, pane.slug, failures);
    await checkSearchRows(page, pane.slug, targets, failures);
    await checkFocusWalk(page, pane.slug, failures);
  } finally {
    await context.close();
  }
  for (const error of pageErrors) {
    failures.push(
      `${pane.slug} | page-level error (outside a tracked click) | ${error}`,
    );
  }
  return { slug: pane.slug, count, failures };
}

async function run() {
  const options = parseArgs(process.argv.slice(2));
  const server = await ensureServer(options.port);
  const browser = await launchBrowser();
  const results = [];
  try {
    for (const pane of options.panes) {
      results.push(await walkPane(browser, server.url, pane));
    }
  } finally {
    await browser.close();
    server.stop();
  }

  let total = 0;
  for (const result of results) {
    for (const failure of result.failures) {
      process.stdout.write(`${failure}\n`);
      total += 1;
    }
    process.stdout.write(
      `${result.failures.length === 0 ? "PASS" : "FAIL"} - ${result.slug}: ${String(result.failures.length)} defect(s) across ${String(result.count)} controls\n`,
    );
  }
  process.stdout.write(
    total === 0
      ? `PASS - 0 defects across ${String(results.reduce((sum, r) => sum + r.count, 0))} controls\n`
      : `FAIL - ${String(total)} defect(s)\n`,
  );
  process.exitCode = total === 0 ? 0 : 1;
}

run().catch((error) => {
  process.stderr.write(`${String(error?.stack ?? error)}\n`);
  process.exitCode = 1;
});
