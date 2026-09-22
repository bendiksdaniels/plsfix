// @vitest-environment jsdom
// Slice H click-through harness: presses every button of the shipped
// taskpane.html + src/main.ts over the STRICT fake Excel host, relay mocked,
// in three shapes - (a) a fresh workbook, (b) a seeded model block, (c) two
// adverse passes (a multi-area selection, a protected sheet) - and asserts
// every press answers like a product should: no crash, some reaction, never
// a raw office.js code or "undefined" standing alone in the toast, and the
// busy latch releasing again. Also runs every PLSFIX_* ribbon command.
// Owns no product logic; a real defect found along the way is catalogued in
// KNOWN_DEFECTS below rather than patched here (src/ is out of scope for
// this slice) and the suite asserts it is STILL there, so a later fix must
// remove the entry.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import type * as RelayModule from "../src/link/relay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "./fakehost";
import { FakeRelay } from "./fakerelay";

enableStrictLoadSemantics();

// The pane builds its own RelayClient from document.baseURI; handed this one
// instead, the same technique test/stress.ppt.support.ts uses for the
// PowerPoint pane (a constructor mock returning an existing instance).
let relay: FakeRelay;
vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return relay;
    },
  };
});

// ---------------------------------------------------------------------------
// Boot (copies test/I.audit.integration.test.ts's pattern - not imported,
// per the brief: that file's helpers are private to its own suite).
// ---------------------------------------------------------------------------

let helpers: FakeHelpers;

function pane(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

// test/fakehost.ts's Office.onReady stub only implements the callback style;
// src/main.ts calls the no-callback, promise-returning style office.js also
// supports, so it is patched here rather than in the shared fake (additive
// only, and this needs an existing line changed).
function patchOnReadyForPromiseStyle(): void {
  const office = (
    globalThis as unknown as {
      Office: {
        onReady: (cb?: (info: { host: string }) => unknown) => unknown;
      };
    }
  ).Office;
  office.onReady = (callback) => {
    const info = { host: "Excel" };
    callback?.(info);
    return Promise.resolve(info);
  };
}

async function drain(rounds = 12): Promise<void> {
  for (let i = 0; i < rounds; i += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

interface BootOptions {
  sheets?: string[];
  seed?: (h: FakeHelpers) => void;
}

async function bootExcel(options: BootOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  // officeKeyStore() (src/link/workspace.ts) falls through to a module-level
  // in-memory Map here: this Node/jsdom combination's global `localStorage`
  // exists but every method on it is undefined (see src/ui/first-run.test.ts
  // for the same finding), so webStorage()'s own probe throws and is caught.
  // vi.resetModules() below re-evaluates that module fresh on the next
  // import, which is what actually gives every boot a clean key store; no
  // manual clearing needed (and `localStorage.clear` is not callable here to
  // even attempt it).
  pane();
  const host = installFakeHost({ sheets: options.sheets ?? ["Model", "Data"] });
  helpers = host.helpers;
  relay = new FakeRelay();
  options.seed?.(helpers);
  patchOnReadyForPromiseStyle();
  await import("../src/main");
  await drain();
}

function findButton(idOrAction: string): HTMLButtonElement | null {
  return (
    (document.getElementById(idOrAction) as HTMLButtonElement | null) ??
    document.querySelector<HTMLButtonElement>(`[data-action="${idOrAction}"]`)
  );
}

async function settle(): Promise<void> {
  await drain();
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

// openShortcutCard() (src/pane/dispatch.ts) falls back to window.open when
// the fake host's Office.context.ui carries no displayDialogAsync - jsdom
// does not implement navigation at all, so the fallback is stubbed out here
// once for the whole file, the same workaround test/I.audit.integration.
// test.ts uses for the same button.
vi.stubGlobal("open", vi.fn());

afterEach(() => {
  uninstallFakeHost();
  vi.restoreAllMocks();
  vi.stubGlobal("open", vi.fn());
});

// ---------------------------------------------------------------------------
// Uncaught error / rejection tracking (the brief: "install window.onerror /
// unhandledrejection hooks and also let vitest's own detection stand").
// Installed once - window persists for the whole file, only document.body is
// swapped per boot - and read by index, never cleared, so a leak from one
// press cannot silently attribute itself to the next.
// ---------------------------------------------------------------------------

const capturedErrors: string[] = [];
window.addEventListener("error", (event) => {
  capturedErrors.push(`error: ${String(event.error ?? event.message)}`);
});
window.addEventListener("unhandledrejection", (event) => {
  capturedErrors.push(`unhandledrejection: ${String(event.reason)}`);
});

// ---------------------------------------------------------------------------
// Reaction fingerprint: a sentinel spliced directly into the toast's own text
// node (never through toast.show(), so the auto-hide timer armed by an
// EARLIER press - which only ever flips #toast's className, never touches
// its children - can never race this into a false pass or fail) plus
// body.innerHTML and the focused element, exactly the three the brief names.
// ---------------------------------------------------------------------------

const SENTINEL = "CLICKTHROUGH-NO-TOAST-YET";

function armSentinel(): void {
  const container = document.getElementById("toast");
  if (!container) return;
  let span = container.querySelector<HTMLElement>(".toast-text");
  if (!span) {
    span = document.createElement("span");
    span.className = "toast-text";
    container.append(span);
  }
  span.textContent = SENTINEL;
}

// ---------------------------------------------------------------------------
// Busy latch: shared.ts's setBusy(busy) disables every [data-action] button,
// a fixed list captured once at module load - never the generated rows, the
// Links tab's own buttons (their own busy scope) or the five with a rule of
// their own the brief names.
// ---------------------------------------------------------------------------

// The brief's five, plus one this suite found the same way: reveal-key
// (src/pane/links-tab.ts wireBoxes(), by its own comment) "is local and
// instant - it touches neither Office nor the store - so it stays out of
// the guard and out of the busy state" and toggles tab.reveal.disabled
// purely off whether a key exists (renderKey), never off busy - forget-key
// legitimately leaves it disabled behind it.
const BUSY_LATCH_EXCEPTIONS = new Set([
  "trace-back",
  "copy-model-check",
  "styles-delete",
  "delete-names",
  "generate-key",
  "reveal-key",
]);

function enabledSnapshot(): Map<string, boolean> {
  const map = new Map<string, boolean>();
  for (const button of document.querySelectorAll<HTMLButtonElement>(
    "button[id], button[data-action]",
  )) {
    const key = button.dataset.action ?? button.id;
    if (BUSY_LATCH_EXCEPTIONS.has(key)) continue;
    map.set(key, !button.disabled);
  }
  return map;
}

function assertBusyReleased(
  pressedKey: string,
  before: Map<string, boolean>,
  after: Map<string, boolean>,
): void {
  for (const [key, wasEnabled] of before) {
    if (!wasEnabled) continue;
    expect(
      after.get(key),
      `after pressing "${pressedKey}", button "${key}" is still disabled (busy latch not released)`,
    ).toBe(true);
  }
}

// ---------------------------------------------------------------------------
// Toast sanity: the product's own sentences, never dispatch()'s fallback,
// never a bare office.js error code standing alone, never a leaked
// "undefined" or a stringified object.
// ---------------------------------------------------------------------------

const BARE_HOST_CODE =
  /^(GeneralException|InvalidArgument|ItemAlreadyExists|ItemNotFound|InvalidOperation|InvalidSelection|UnsupportedOperation|AccessDenied|PropertyNotLoaded)$/;

function assertToastSanity(key: string, toast: string): void {
  expect(
    toast,
    `"${key}" toast fell through dispatch()'s "Unknown action" case`,
  ).not.toMatch(/^Unknown action/);
  expect(
    toast,
    `"${key}" toast is a bare office.js code with no sentence around it`,
  ).not.toMatch(BARE_HOST_CODE);
  expect(
    toast,
    `"${key}" toast leaked the literal text "undefined"`,
  ).not.toMatch(/undefined/);
  expect(toast, `"${key}" toast leaked a stringified object`).not.toMatch(
    /\[object/,
  );
}

// ---------------------------------------------------------------------------
// KNOWN_DEFECTS: a real product behaviour this slice found and may not fix
// (src/ is out of scope here). Each entry's predicate must still match, or
// the suite fails - the signal a later fix slice removes the entry against.
// ---------------------------------------------------------------------------

interface KnownDefect {
  description: string;
  cause: string;
  stillPresent: (toast: string, reacted: boolean) => boolean;
}

const KNOWN_DEFECTS: Record<string, KnownDefect> = {};

// Two icon buttons whose whole job is re-running a read something earlier in
// THIS SAME pass already ran once, with nothing in between to make a second
// run answer differently - correctly a no-op, not a dead button, the same
// shape as re-clicking the already-active tab below:
//  - refresh-selection re-reads exactly what boot()'s own eager
//    refreshSelection() call already displayed (src/main.ts's boot():
//    "if (excelConnected) await refreshSelection();"); every seed this
//    suite ever gives a boot is set up BEFORE that import, so its first
//    press always meets a selection that has not moved since;
//  - refresh-sheets duplicates tab-workbook's own click handler (also
//    `refreshSheets()`, src/main.ts wireControls()), and the static-button
//    pass below reaches every tab before it reaches this icon.
const IDEMPOTENT_UNTIL_SELECTION_MOVES = new Set([
  "refresh-selection",
  "refresh-sheets",
]);

// ---------------------------------------------------------------------------
// The one generic press: fires a real click, waits it out, then holds it to
// every rule every press owes regardless of which button it was. Takes the
// element directly so a generated row (no id, no data-action) can share the
// exact same checks as a static button; press(key) below is the common case.
// ---------------------------------------------------------------------------

async function pressElement(
  button: HTMLButtonElement,
  label: string,
): Promise<string> {
  // A disabled button (trace-back with an empty back stack, copy-model-check
  // before a report exists) correctly does nothing on a real click - jsdom
  // honours the disabled attribute the same as a browser, so nothing fires.
  // Pressed all the same, for coverage, just not held to reacting.
  if (button.disabled) {
    expect(
      () => button.click(),
      `"${label}" threw synchronously on click`,
    ).not.toThrow();
    await settle();
    return toastText();
  }
  // Re-pressing the tab that is already active is correctly a no-op (same
  // rule scripts/ux/sweep-probe.js applies): it proves the tab stays
  // consistent, not that it reacts.
  const alreadyActiveTab =
    button.getAttribute("role") === "tab" &&
    button.getAttribute("aria-selected") === "true";
  const expectNoNewReaction =
    alreadyActiveTab || IDEMPOTENT_UNTIL_SELECTION_MOVES.has(label);
  const busyBefore = enabledSnapshot();
  armSentinel();
  const htmlBefore = document.body.innerHTML;
  const activeBefore = document.activeElement;
  const errorsBefore = capturedErrors.length;

  expect(
    () => button.click(),
    `"${label}" threw synchronously on click`,
  ).not.toThrow();
  await settle();

  expect(
    capturedErrors.slice(errorsBefore),
    `"${label}" raised an uncaught error or unhandled rejection`,
  ).toEqual([]);

  const toast = toastText();
  const htmlAfter = document.body.innerHTML;
  const activeAfter = document.activeElement;
  const reacted =
    toast !== SENTINEL ||
    htmlAfter !== htmlBefore ||
    activeAfter !== activeBefore;

  const defect = KNOWN_DEFECTS[label];
  if (defect) {
    expect(
      defect.stillPresent(toast, reacted),
      `KNOWN_DEFECTS["${label}"] (${defect.description}) no longer reproduces - remove the entry`,
    ).toBe(true);
    return toast;
  }

  expect(
    reacted || expectNoNewReaction,
    `button "${label}" produced no reaction at all (dead button)`,
  ).toBe(true);
  assertToastSanity(label, toast);
  assertBusyReleased(label, busyBefore, enabledSnapshot());
  return toast;
}

async function press(key: string): Promise<string> {
  const button = findButton(key);
  expect(button, `coverage: no button found for "${key}"`).toBeTruthy();
  return pressElement(button!, key);
}

// ---------------------------------------------------------------------------
// Static coverage: every <button> the shipped taskpane.html carries, keyed
// the same way dispatch()/help.ts do (data-action first, else id). Logo
// swatches are excluded by construction - src/pane/brand-tab.ts only ever
// builds them from a decoded image FILE (extractLogoColors, canvas-based
// colour extraction), which jsdom cannot decode; the shipped markup ships
// #logo-swatches empty, so no <button> lives there to find in the first
// place. Skipped here, deliberately, with this comment as the record.
// ---------------------------------------------------------------------------

function staticButtonKeys(): string[] {
  const html = readFileSync(join(process.cwd(), "taskpane.html"), "utf8");
  // DOMParser (standard Web API, part of every jsdom environment) reads the
  // shipped markup into its own detached document, without booting it or
  // touching the live `document` - no need for the "jsdom" package's own
  // JSDOM class (not a project dependency's public API elsewhere in this
  // repo, and untyped here).
  const scratch = new DOMParser().parseFromString(html, "text/html");
  const keys: string[] = [];
  for (const button of scratch.querySelectorAll("button")) {
    const key = button.getAttribute("data-action") ?? button.id;
    if (!key) {
      throw new Error(
        `coverage: a taskpane.html button has neither id nor data-action: ${button.outerHTML.slice(0, 80)}`,
      );
    }
    keys.push(key);
  }
  return keys;
}

const pressedStatic = new Set<string>();

async function pressStatic(key: string): Promise<string> {
  pressedStatic.add(key);
  return press(key);
}

type GeneratedKind =
  | "sheet-name-button"
  | "sheet-eye-button"
  | "link-row-tick"
  | "model-check-finding-row"
  | "trace-chip"
  | "find-result-row";

const generatedCovered = new Set<GeneratedKind>();

// ---------------------------------------------------------------------------
// Seeding: the model block state (b)/(c) share - labels in column A, numbers
// and formulas in B:D over six rows, one hardcoded growth rate per row (a
// model-check hardcodeInFormula finding) and D1's precedent (C1) wired for
// Smart Track.
// ---------------------------------------------------------------------------

function seedModelBlock(h: FakeHelpers): void {
  h.seed("Model!A1", [
    ["Revenue", 100, 110, { formula: "=C1*1.07" }],
    ["COGS", 40, 44, { formula: "=C2*1.1" }],
    [
      "Gross profit",
      { formula: "=B1-B2" },
      { formula: "=C1-C2" },
      { formula: "=D1-D2" },
    ],
    ["Opex", 20, 22, { formula: "=C4*1.05" }],
    [
      "EBITDA",
      { formula: "=B3-B4" },
      { formula: "=C3-C4" },
      { formula: "=D3-D4" },
    ],
    [
      "Margin",
      { formula: "=B5/B1" },
      { formula: "=C5/C1" },
      { formula: "=D5/D1" },
    ],
  ]);
  h.addChart("Model", {
    name: "Revenue bridge",
    width: 400,
    height: 200,
    chartType: "ColumnClustered",
  });
  // An unused custom style, for the Workbook tab's style scrubber.
  h.addStyle("StaleAccent", false);
  // A broken defined name (Excel's own #REF! rewrite), for the Workbook
  // tab's name scrubber.
  h.addName("CostBase", "=Model!$B$1");
  h.breakName("CostBase");
  h.select("Model!A1:D6");
  primeTraceD1(h);
}

// D1's precedent (C1), for Smart Track - re-armed by name wherever a test
// needs it, since an earlier tool in the same pass may otherwise move the
// active cell the fake's getDirectPrecedents() reads.
function primeTraceD1(h: FakeHelpers): void {
  h.setActiveCell("Model!D1");
  h.setPrecedents("Model!D1", [{ address: "Model!C1", cellCount: 1 }]);
}

// A link key and one exported range, driven entirely through the pane's own
// buttons (Links tab), so the Links list gets a real row with a tick box.
async function seedLinkAndRow(): Promise<void> {
  const linksTab = findButton("tab-links")!;
  linksTab.click();
  await settle();
  await pressStatic("generate-key");
  helpers.select("Model!A1:D6");
  await pressStatic("export-selection");
}

function tickFirstLinkRow(): void {
  const box = document.querySelector<HTMLInputElement>(
    "#workbook-links tr[data-link-id] input[type=checkbox]",
  );
  expect(box, "seed: no link row to tick").toBeTruthy();
  box!.checked = true;
  box!.dispatchEvent(new Event("change"));
}

// ---------------------------------------------------------------------------
// State (a): fresh workbook, active cell A1, nothing seeded.
// ---------------------------------------------------------------------------

describe("state (a): fresh workbook, active cell A1, nothing seeded", () => {
  it("presses every static button and every ribbon command once", async () => {
    await bootExcel();

    for (const key of new Set(staticButtonKeys())) {
      // project-ok/project-cancel need their prompt open first; new-project
      // opens it, and sits right before them in document order, so this
      // single natural pass already covers the precondition.
      await pressStatic(key);
    }

    // Ribbon: every PLSFIX_* command, fired the way Office's shared runtime
    // does - a bare `event.completed` callback, nothing else - resolves and
    // completes without throwing, before Excel has ever connected in this
    // pass or not (state (a) still has excelConnected true here: the fake
    // host always answers Office.onReady with Excel; "not connected" is a
    // different, host-rejection scenario covered by test/I.audit.*).
    for (const [id, handler] of helpers.actions()) {
      const completed = vi.fn();
      expect(
        () => handler({ completed }),
        `ribbon "${id}" threw synchronously`,
      ).not.toThrow();
      await settle();
      expect(
        completed,
        `ribbon "${id}" never called event.completed()`,
      ).toHaveBeenCalled();
    }
  });

  it("answers the brief's specific sentences with nothing selected", async () => {
    await bootExcel();

    expect(await press("undo")).toBe("There is no pls,fix action to undo yet.");
    expect(await press("paste-values")).toBe("Mark a copy source first.");
    expect(await press("template-dcf")).toMatch(/^Template written:/);

    const linksTab = findButton("tab-links")!;
    linksTab.click();
    await settle();
    expect(await press("export-selection")).toBe(
      "Generate a link key first (Links > Link key).",
    );
    for (const id of [
      "push-selected",
      "go-to-source",
      "remove-link",
      "move-to-project",
    ]) {
      expect(await press(id)).toBe("Select a link in the list first.");
    }

    findButton("new-project")!.click();
    await settle();
    expect(await press("project-ok")).toBe("Type a project name.");
    expect(await press("project-cancel")).toBe("Cancelled.");
  });
});

// ---------------------------------------------------------------------------
// State (b): a seeded model block selected, a chart on the sheet, a second
// sheet, a link, a link key, an unused style and a model-check finding.
// ---------------------------------------------------------------------------

describe("state (b): a seeded model block, a chart, a second sheet", () => {
  it("presses every static button again, plus every generated row and chip", async () => {
    await bootExcel({ seed: seedModelBlock });
    (document.getElementById("find-query") as HTMLInputElement).value =
      "Revenue";

    await seedLinkAndRow();
    tickFirstLinkRow();

    // Ninety buttons fired in sequence over one small block is not gentle:
    // several tools (template inserts, Smart Painter, anything that
    // `.select()`s its own result) move the selection and/or the active
    // cell as a side effect, and insert-color-key rewrote A1 itself in an
    // early run of this suite. trace-precedents and find both depend on
    // exact cell content or the active cell at the moment they run, so both
    // are held out of the generic pass and re-armed immediately before
    // their own explicit press, the only way to make either deterministic
    // against however the ninety before it left the sheet.
    const keys = new Set(staticButtonKeys());
    keys.delete("trace-precedents");
    keys.delete("find");
    for (const key of keys) {
      await pressStatic(key);
    }
    primeTraceD1(helpers);
    findButton("tab-tools")!.click();
    await settle();
    await pressStatic("trace-precedents");

    helpers.seed("Model!A1", [["Revenue"]]);
    findButton("tab-workbook")!.click();
    await settle();
    await pressStatic("find");

    // Generated: the Workbook tab's sheet explorer - one name button and one
    // eye button per visible sheet (src/pane/workbook-tab.ts sheetRow()).
    findButton("tab-workbook")!.click();
    await settle();
    const sheetRows = [
      ...document.querySelectorAll<HTMLDivElement>("#sheet-list .sheet-row"),
    ];
    expect(
      sheetRows.length,
      "seed: expected two sheets in the explorer",
    ).toBeGreaterThanOrEqual(2);
    for (const row of sheetRows) {
      const nameButton =
        row.querySelector<HTMLButtonElement>("button.sheet-name");
      if (nameButton) {
        generatedCovered.add("sheet-name-button");
        await pressElement(
          nameButton,
          `sheet-name:${nameButton.textContent ?? ""}`,
        );
      }
      const eyeButton = row.querySelector<HTMLButtonElement>("button.eye");
      if (eyeButton) {
        generatedCovered.add("sheet-eye-button");
        await pressElement(eyeButton, `sheet-eye:${eyeButton.title}`);
      }
    }

    // Generated: the Links tab's tick box - already exercised above
    // (tickFirstLinkRow), counted here for the coverage census.
    generatedCovered.add("link-row-tick");

    // Generated: the model-check report's finding rows (run-model-check ran
    // as part of the static pass above, against the seeded hardcoded-growth
    // formulas in column D).
    findButton("tab-workbook")!.click();
    await settle();
    const findingRows = [
      ...document.querySelectorAll<HTMLButtonElement>(
        "#model-check-list > button",
      ),
    ];
    expect(
      findingRows.length,
      "seed: expected at least one model-check finding",
    ).toBeGreaterThan(0);
    for (const row of findingRows) {
      generatedCovered.add("model-check-finding-row");
      await pressElement(row, `model-check-row:${row.title}`);
    }

    // Generated: Super Find's result rows, from the explicit find press above.
    const findRows = [
      ...document.querySelectorAll<HTMLButtonElement>("#find-results > button"),
    ];
    expect(
      findRows.length,
      'seed: expected at least one find hit for "Revenue"',
    ).toBeGreaterThan(0);
    for (const row of findRows) {
      generatedCovered.add("find-result-row");
      await pressElement(row, `find-row:${row.title}`);
    }

    // Generated: Smart Track's chips, from the trace-precedents press above.
    findButton("tab-tools")!.click();
    await settle();
    const chips = [
      ...document.querySelectorAll<HTMLButtonElement>(
        "#trace-chips button.chip",
      ),
    ];
    expect(
      chips.length,
      "seed: expected at least one precedent chip",
    ).toBeGreaterThan(0);
    for (const chip of chips) {
      generatedCovered.add("trace-chip");
      await pressElement(chip, `trace-chip:${chip.textContent ?? ""}`);
    }

    // Ribbon again, now against a workbook with real content.
    for (const [id, handler] of helpers.actions()) {
      const completed = vi.fn();
      expect(
        () => handler({ completed }),
        `ribbon "${id}" threw synchronously`,
      ).not.toThrow();
      await settle();
      expect(
        completed,
        `ribbon "${id}" never called event.completed()`,
      ).toHaveBeenCalled();
    }
  });
});

// ---------------------------------------------------------------------------
// State (c): adverse - a multi-area selection, then (a fresh boot, second
// pass) a protected sheet. Both keep the seeded model block; neither repeats
// the generated-content census state (b) already covered.
// ---------------------------------------------------------------------------

const SELECT_SINGLE_RANGE = /: select a single range$/;
const SHEET_PROTECTED = /: this sheet is protected, nothing was changed$/;

describe("state (c): adverse - a multi-area selection", () => {
  it("presses every static button and shows single-range tools their own refusal", async () => {
    await bootExcel({ seed: seedModelBlock });
    helpers.selectAreas(["Model!A1:B2", "Model!D4:D6"]);

    const keys = new Set(staticButtonKeys());
    keys.delete("trace-precedents");
    keys.delete("find");
    const seenSingleRangeRefusal: string[] = [];
    for (const key of keys) {
      const toast = await pressStatic(key);
      if (SELECT_SINGLE_RANGE.test(toast)) seenSingleRangeRefusal.push(key);
      // A multi-area selection must never desync the areas back to one: the
      // adverse condition holds for every remaining press in this pass.
      helpers.selectAreas(["Model!A1:B2", "Model!D4:D6"]);
    }
    primeTraceD1(helpers);
    helpers.selectAreas(["Model!A1:B2", "Model!D4:D6"]);
    findButton("tab-tools")!.click();
    await settle();
    await pressStatic("trace-precedents");

    helpers.seed("Model!A1", [["Revenue"]]);
    findButton("tab-workbook")!.click();
    await settle();
    await pressStatic("find");

    // Proof the mechanism actually engaged this pass, not just that nothing
    // crashed - selectedSingleRange's own refusal (src/excel/internal.ts),
    // worded per tool by its own `stage`.
    expect(
      seenSingleRangeRefusal,
      "no single-range tool refused the multi-area selection this pass",
    ).not.toEqual([]);
  });
});

describe("state (c): adverse - a protected sheet", () => {
  it("presses every static button and shows writers their own refusal", async () => {
    await bootExcel({ seed: seedModelBlock });
    helpers.protectSheet("Model", []);

    const keys = new Set(staticButtonKeys());
    keys.delete("trace-precedents");
    keys.delete("find");
    const seenProtectedRefusal: string[] = [];
    for (const key of keys) {
      const toast = await pressStatic(key);
      if (SHEET_PROTECTED.test(toast)) seenProtectedRefusal.push(key);
    }
    primeTraceD1(helpers);
    findButton("tab-tools")!.click();
    await settle();
    await pressStatic("trace-precedents");

    findButton("tab-workbook")!.click();
    await settle();
    await pressStatic("find");

    expect(
      seenProtectedRefusal,
      "no writer refused the protected sheet this pass",
    ).not.toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// Two-click buttons: the first press only arms; the second is the real one.
// ---------------------------------------------------------------------------

describe("two-click confirms", () => {
  it("styles-delete arms on the first click and deletes on the second", async () => {
    await bootExcel({ seed: seedModelBlock });
    findButton("tab-workbook")!.click();
    await settle();
    await press("styles-scan");

    const button = findButton("styles-delete")!;
    expect(button.classList.contains("armed")).toBe(false);
    await press("styles-delete");
    expect(
      button.classList.contains("armed"),
      "first press should only arm it",
    ).toBe(true);

    const secondToast = await press("styles-delete");
    expect(secondToast).toMatch(/^Deleted \d+ unused styles?$/);
    expect(
      button.classList.contains("armed"),
      "the confirmed delete disarms it again",
    ).toBe(false);
  });

  it("delete-names arms on the first click and deletes on the second", async () => {
    await bootExcel({ seed: seedModelBlock });
    findButton("tab-workbook")!.click();
    await settle();
    await press("scan-names");

    const button = findButton("delete-names")!;
    expect(button.classList.contains("armed")).toBe(false);
    await press("delete-names");
    expect(
      button.classList.contains("armed"),
      "first press should only arm it",
    ).toBe(true);

    const secondToast = await press("delete-names");
    expect(secondToast).toMatch(/^Deleted \d+ broken names?$/);
    expect(
      button.classList.contains("armed"),
      "the confirmed delete disarms it again",
    ).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Coverage: every static button the shipped taskpane.html carries, and every
// documented generated-content kind, was pressed by at least one state/pass
// above. A button added to the markup later fails this until it is pressed.
// ---------------------------------------------------------------------------

describe("coverage", () => {
  it("pressed every static button at least once", () => {
    const required = new Set(staticButtonKeys());
    const missing = [...required].filter((key) => !pressedStatic.has(key));
    expect(missing, "static buttons never pressed by any state").toEqual([]);
  });

  it("exercised every documented generated-content kind at least once", () => {
    const required: GeneratedKind[] = [
      "sheet-name-button",
      "sheet-eye-button",
      "link-row-tick",
      "model-check-finding-row",
      "trace-chip",
      "find-result-row",
    ];
    const missing = required.filter((kind) => !generatedCovered.has(kind));
    expect(missing, "generated-content kinds never exercised").toEqual([]);
  });
});
