// @vitest-environment jsdom
// Audit coverage for the Excel ribbon/shortcut command table: at the base
// this file ran in zero tests (0% coverage), so registerCommands()'s wiring,
// its success/failure promise chain and PLSFIX_SHOWPANE had never executed.
// Every dependency is mocked (this is a wiring test, not a re-test of what
// each action does - those have their own tests under their owning slices)
// so the table can be driven purely through a stubbed Office.actions global.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ADDIN } from "../../manifest/spec";

vi.mock("../excel", () => ({
  addCagrLabel: vi.fn(async () => undefined),
  applyBorderCycle: vi.fn(async () => undefined),
  applyColumnWidthCycle: vi.fn(async () => undefined),
  applyDecimalStep: vi.fn(async () => undefined),
  applyFillCycle: vi.fn(async () => undefined),
  applyFontColorCycle: vi.fn(async () => undefined),
  applyNumberCycle: vi.fn(async () => undefined),
  applyRowHeightCycle: vi.fn(async () => undefined),
  applyRowStyleCycle: vi.fn(async () => undefined),
  applySignFlip: vi.fn(async () => undefined),
  autocolorSelection: vi.fn(async () => undefined),
  fastFillAuto: vi.fn(async () => undefined),
  formatSelectedChart: vi.fn(async () => undefined),
  insertCagr: vi.fn(async () => undefined),
  insertConsistentRounding: vi.fn(async () => undefined),
  insertTornado: vi.fn(async () => undefined),
  insertWaterfall: vi.fn(async () => undefined),
  lastUndoSkipped: vi.fn(() => undefined),
  markCopySource: vi.fn(async () => undefined),
  pasteSpecial: vi.fn(async () => undefined),
  pastePreserveFormulas: vi.fn(async () => undefined),
  scaleSelection: vi.fn(async () => undefined),
  toggleIfErrorGuard: vi.fn(async () => undefined),
  undoLastAction: vi.fn(async () => undefined),
  unpivotSelection: vi.fn(async () => undefined),
}));
vi.mock("./find-panel", () => ({ focusFind: vi.fn(async () => undefined) }));
vi.mock("./paint-slots", () => ({
  capturePaintSlot: vi.fn(async () => undefined),
  applyPaintSlot: vi.fn(async () => undefined),
}));
vi.mock("./share-panel", () => ({
  prepareShare: vi.fn(async () => undefined),
}));
vi.mock("./shared", () => ({
  APP_VERSION: "v9.9.999",
  refreshSelection: vi.fn(async () => undefined),
  renderActionState: vi.fn(() => undefined),
  toast: { show: vi.fn() },
}));
vi.mock("./styles-panel", () => ({
  focusStyles: vi.fn(async () => undefined),
}));
vi.mock("./trace-panel", () => ({
  startTrace: vi.fn(async () => undefined),
  toggleAudit: vi.fn(async () => undefined),
}));
vi.mock("./workbook-tab", () => ({
  insertTocSheet: vi.fn(async () => undefined),
}));

import {
  addCagrLabel,
  applyBorderCycle,
  applyColumnWidthCycle,
  applyDecimalStep,
  applyFillCycle,
  applyFontColorCycle,
  applyNumberCycle,
  applyRowHeightCycle,
  applyRowStyleCycle,
  applySignFlip,
  autocolorSelection,
  fastFillAuto,
  formatSelectedChart,
  insertCagr,
  insertConsistentRounding,
  insertTornado,
  insertWaterfall,
  lastUndoSkipped,
  markCopySource,
  pasteSpecial,
  pastePreserveFormulas,
  scaleSelection,
  toggleIfErrorGuard,
  undoLastAction,
  unpivotSelection,
} from "../excel";
import { focusFind } from "./find-panel";
import { capturePaintSlot, applyPaintSlot } from "./paint-slots";
import { prepareShare } from "./share-panel";
import { refreshSelection, renderActionState, toast } from "./shared";
import { focusStyles } from "./styles-panel";
import { startTrace, toggleAudit } from "./trace-panel";
import { insertTocSheet } from "./workbook-tab";
import { registerCommands } from "./commands";

type Handler = (event?: { completed: () => void }) => void;

function stubOffice(): Map<string, Handler> {
  const associated = new Map<string, Handler>();
  (globalThis as { Office?: unknown }).Office = {
    actions: {
      associate: (id: string, handler: Handler) => associated.set(id, handler),
    },
    addin: { showAsTaskpane: vi.fn(async () => undefined) },
  };
  return associated;
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

// Every case: the FunctionName, the mock it must reach, and the exact
// arguments it must be called with (undefined means "called with no args").
const CASES: [string, () => unknown, unknown[]][] = [
  ["PLSFIX_AUTOCOLOR", () => autocolorSelection, []],
  ["PLSFIX_AUDIT", () => toggleAudit, []],
  ["PLSFIX_TRACE_PRE", () => startTrace, ["precedents", true]],
  ["PLSFIX_TRACE_DEP", () => startTrace, ["dependents", true]],
  ["PLSFIX_FILLRIGHT", () => fastFillAuto, ["right"]],
  ["PLSFIX_FILLDOWN", () => fastFillAuto, ["down"]],
  ["PLSFIX_IFERROR", () => toggleIfErrorGuard, []],
  ["PLSFIX_SCALEUP", () => scaleSelection, [1000]],
  ["PLSFIX_SCALEDOWN", () => scaleSelection, [0.001]],
  ["PLSFIX_UNDO", () => undoLastAction, []],
  ["PLSFIX_COPYSRC", () => markCopySource, []],
  ["PLSFIX_PASTE_VALUES", () => pasteSpecial, ["values"]],
  ["PLSFIX_PASTE_FORMATS", () => pasteSpecial, ["formats"]],
  ["PLSFIX_PASTE_EXACT", () => pastePreserveFormulas, []],
  ["PLSFIX_PASTE_TRANSPOSE", () => pasteSpecial, ["transpose"]],
  ["PLSFIX_CAGR", () => insertCagr, []],
  ["PLSFIX_ROUND", () => insertConsistentRounding, []],
  ["PLSFIX_SIGN", () => applySignFlip, []],
  ["PLSFIX_DEC_MORE", () => applyDecimalStep, [1]],
  ["PLSFIX_DEC_LESS", () => applyDecimalStep, [-1]],
  ["PLSFIX_CYC_GENERAL", () => applyNumberCycle, ["general"]],
  ["PLSFIX_CYC_DATE", () => applyNumberCycle, ["date"]],
  ["PLSFIX_CYC_CURRENCY", () => applyNumberCycle, ["currency"]],
  ["PLSFIX_CYC_PERCENT", () => applyNumberCycle, ["percent"]],
  ["PLSFIX_CYC_MULTIPLE", () => applyNumberCycle, ["multiple"]],
  ["PLSFIX_CYC_TITLE", () => applyRowStyleCycle, ["title"]],
  ["PLSFIX_CYC_RESULT", () => applyRowStyleCycle, ["result"]],
  ["PLSFIX_CYC_ITEM", () => applyRowStyleCycle, ["item"]],
  ["PLSFIX_CYC_FILL", () => applyFillCycle, []],
  ["PLSFIX_CYC_FONT", () => applyFontColorCycle, []],
  ["PLSFIX_CYC_BORDER", () => applyBorderCycle, []],
  ["PLSFIX_CYC_ROWH", () => applyRowHeightCycle, []],
  ["PLSFIX_CYC_COLW", () => applyColumnWidthCycle, []],
  ["PLSFIX_PAINT_CAP1", () => capturePaintSlot, [1]],
  ["PLSFIX_PAINT_CAP2", () => capturePaintSlot, [2]],
  ["PLSFIX_PAINT_CAP3", () => capturePaintSlot, [3]],
  ["PLSFIX_PAINT_APP1", () => applyPaintSlot, [1]],
  ["PLSFIX_PAINT_APP2", () => applyPaintSlot, [2]],
  ["PLSFIX_PAINT_APP3", () => applyPaintSlot, [3]],
  ["PLSFIX_WATERFALL", () => insertWaterfall, []],
  ["PLSFIX_TORNADO", () => insertTornado, []],
  ["PLSFIX_CHARTFMT", () => formatSelectedChart, []],
  ["PLSFIX_CHART_CAGR", () => addCagrLabel, []],
  ["PLSFIX_UNPIVOT", () => unpivotSelection, []],
  ["PLSFIX_TOC", () => insertTocSheet, []],
  ["PLSFIX_SHARE", () => prepareShare, []],
  ["PLSFIX_FIND", () => focusFind, []],
  ["PLSFIX_STYLES_SCAN", () => focusStyles, []],
];

// Ids the table registers on purpose with no ribbon button and no keyboard
// shortcut anywhere in manifest/spec.ts or public/shortcuts.json: reachable
// only from inside the pane's own Tools/Workbook tabs today (confirmed
// against src/pane/dispatch.ts). Reported to Daniel as a completeness gap,
// not fixed here - public/shortcuts.json is off limits to every slice.
const NO_SHORTCUT_YET = [
  "PLSFIX_ROUND",
  "PLSFIX_PAINT_CAP1",
  "PLSFIX_PAINT_CAP2",
  "PLSFIX_PAINT_CAP3",
  "PLSFIX_TORNADO",
  "PLSFIX_UNPIVOT",
  "PLSFIX_SHARE",
];

// import.meta.url is not a file URL under the jsdom environment (see
// src/pane/links-tab.test.ts), so this reads from the repo root instead.
function shortcutsJsonIds(): string[] {
  const json = JSON.parse(
    readFileSync(join(process.cwd(), "public/shortcuts.json"), "utf8"),
  ) as { actions: { id: string }[] };
  return json.actions.map((action) => action.id);
}

function manifestRibbonIds(): string[] {
  const workbook = ADDIN.hosts.find((host) => host.name === "Workbook")!;
  return workbook.groups
    .flatMap((group) => group.buttons)
    .filter((button) => button.action.kind === "function")
    .map((button) => (button.action as { name: string }).name);
}

beforeEach(() => {
  vi.clearAllMocks();
});
afterEach(() => {
  delete (globalThis as { Office?: unknown }).Office;
});

describe("wiring: every PLSFIX_* id reaches its handler from the table alone", () => {
  it("registers exactly the ids in CASES plus PLSFIX_SHOWPANE, nothing else", () => {
    const associated = stubOffice();
    registerCommands();
    expect([...associated.keys()].sort()).toEqual(
      [...CASES.map(([id]) => id), "PLSFIX_SHOWPANE"].sort(),
    );
  });

  it.each(CASES)(
    "%s calls its own action with its own arguments",
    async (id, getMock, args) => {
      const associated = stubOffice();
      registerCommands();
      const completed = vi.fn();
      associated.get(id)!({ completed });
      await settle();

      const mock = getMock() as ReturnType<typeof vi.fn>;
      expect(mock).toHaveBeenCalledTimes(1);
      expect(mock).toHaveBeenCalledWith(...args);
      expect(refreshSelection).toHaveBeenCalledOnce();
      expect(toast.show).not.toHaveBeenCalled();
      expect(lastUndoSkipped).toHaveBeenCalledOnce();
      expect(renderActionState).toHaveBeenCalledOnce();
      expect(completed).toHaveBeenCalledOnce();
    },
  );

  it("every id public/shortcuts.json declares has a handler in the table", () => {
    const table = new Set(CASES.map(([id]) => id));
    table.add("PLSFIX_SHOWPANE");
    for (const id of shortcutsJsonIds()) {
      expect(table.has(id)).toBe(true);
    }
  });

  it("every ribbon FunctionName the Workbook host declares has a handler in the table", () => {
    const table = new Set(CASES.map(([id]) => id));
    for (const name of manifestRibbonIds()) {
      expect(table.has(name)).toBe(true);
    }
  });

  it("flags every table entry that reaches no ribbon button and no shortcut key yet", () => {
    // A regression guard, not a design requirement: if this list grows, either
    // wire the new id into public/shortcuts.json (or a ribbon button) or add
    // it here on purpose. See NO_SHORTCUT_YET's comment for why the current
    // seven are pane-button-only today.
    const reachable = new Set([...shortcutsJsonIds(), ...manifestRibbonIds()]);
    const orphaned = CASES.map(([id]) => id).filter((id) => !reachable.has(id));
    expect(orphaned.sort()).toEqual([...NO_SHORTCUT_YET].sort());
  });
});

describe("the promise chain when a command fails", () => {
  it("toasts the error named by its FunctionName, still refreshes and completes", async () => {
    const associated = stubOffice();
    vi.mocked(autocolorSelection).mockRejectedValueOnce(new Error("boom"));
    registerCommands();
    const completed = vi.fn();
    associated.get("PLSFIX_AUTOCOLOR")!({ completed });
    await settle();

    expect(toast.show).toHaveBeenCalledWith(
      "boom",
      "error",
      expect.stringContaining("PLSFIX_AUTOCOLOR"),
    );
    expect(lastUndoSkipped).toHaveBeenCalledOnce();
    expect(renderActionState).toHaveBeenCalledOnce();
    expect(completed).toHaveBeenCalledOnce();
  });

  it("shows the pane before toasting, so a closed pane's toast is not lost", async () => {
    const associated = stubOffice();
    vi.mocked(pasteSpecial).mockRejectedValueOnce(
      new Error("Mark a copy source first."),
    );
    registerCommands();
    const completed = vi.fn();
    associated.get("PLSFIX_PASTE_VALUES")!({ completed });
    await settle();

    const office = (
      globalThis as { Office: { addin: { showAsTaskpane: () => void } } }
    ).Office;
    expect(office.addin.showAsTaskpane).toHaveBeenCalledOnce();
    expect(toast.show).toHaveBeenCalledWith(
      "Mark a copy source first.",
      "error",
      expect.stringContaining("PLSFIX_PASTE_VALUES"),
    );
    expect(completed).toHaveBeenCalledOnce();
  });

  it("does not show the pane when a command succeeds", async () => {
    const associated = stubOffice();
    registerCommands();
    const completed = vi.fn();
    associated.get("PLSFIX_PASTE_VALUES")!({ completed });
    await settle();

    const office = (
      globalThis as { Office: { addin: { showAsTaskpane: () => void } } }
    ).Office;
    expect(office.addin.showAsTaskpane).not.toHaveBeenCalled();
    expect(completed).toHaveBeenCalledOnce();
  });
});

describe("PLSFIX_SHOWPANE", () => {
  it("shows the pane as a task pane and completes the event", async () => {
    const associated = stubOffice();
    registerCommands();
    const completed = vi.fn();
    associated.get("PLSFIX_SHOWPANE")!({ completed });
    await settle();

    const office = (
      globalThis as { Office: { addin: { showAsTaskpane: () => void } } }
    ).Office;
    expect(office.addin.showAsTaskpane).toHaveBeenCalledOnce();
    expect(completed).toHaveBeenCalledOnce();
  });

  it("completes even when Office.addin is missing", async () => {
    const associated = new Map<string, Handler>();
    (globalThis as { Office?: unknown }).Office = {
      actions: {
        associate: (id: string, handler: Handler) =>
          associated.set(id, handler),
      },
    };
    registerCommands();
    const completed = vi.fn();
    expect(() =>
      associated.get("PLSFIX_SHOWPANE")!({ completed }),
    ).not.toThrow();
    await settle();
    expect(completed).toHaveBeenCalledOnce();
  });

  it("completes even when showAsTaskpane rejects", async () => {
    const associated = stubOffice();
    const office = (
      globalThis as { Office: { addin: { showAsTaskpane: () => void } } }
    ).Office;
    vi.mocked(office.addin.showAsTaskpane).mockRejectedValueOnce(
      new Error("no host"),
    );
    registerCommands();
    const completed = vi.fn();
    expect(() =>
      associated.get("PLSFIX_SHOWPANE")!({ completed }),
    ).not.toThrow();
    await settle();
    expect(completed).toHaveBeenCalledOnce();
  });
});

describe("a host without Office.actions", () => {
  it("does nothing, never throws", () => {
    (globalThis as { Office?: unknown }).Office = {};
    expect(() => registerCommands()).not.toThrow();
  });
});
