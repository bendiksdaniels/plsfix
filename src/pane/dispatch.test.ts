// @vitest-environment jsdom
// dispatch() has never run in a test (0% coverage). This drives every
// data-action taskpane.html actually ships, plus the bespoke styles-delete
// confirm, through the real routing logic with ../excel and the pane panels
// mocked - no Office host needed. Regression coverage for a shadowing bug a
// sibling slice found: a prefix branch (action.startsWith("cycle-row-"))
// swallowed cycle-row-height before its own exact switch case ever ran.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  applyColumnWidthCycle,
  applyPinstripes,
  applyRowHeightCycle,
  applyRowStyleCycle,
  insertCompsStats,
  insertFootballField,
} from "../excel";
import { isExcelReady } from "./shared";
import { dispatch } from "./dispatch";

vi.mock("../excel", () => ({
  addCagrLabel: vi.fn(async () => "cagr label"),
  applyBorderCycle: vi.fn(async () => undefined),
  applyColumnWidthCycle: vi.fn(async () => undefined),
  applyDecimalStep: vi.fn(async () => undefined),
  applyFillCycle: vi.fn(async () => undefined),
  applyFontColorCycle: vi.fn(async () => undefined),
  applyNumberCycle: vi.fn(async () => undefined),
  applyNumberFormat: vi.fn(async () => undefined),
  applyPinstripes: vi.fn(async () => "pinstripes ok"),
  applyPreset: vi.fn(async () => undefined),
  applyRowHeightCycle: vi.fn(async () => undefined),
  applyRowStyleCycle: vi.fn(async () => undefined),
  applySignFlip: vi.fn(async () => undefined),
  autocolorSelection: vi.fn(async () => "autocolor ok"),
  clearFormats: vi.fn(async () => undefined),
  fastFillAuto: vi.fn(async () => undefined),
  formatSelectedChart: vi.fn(async () => undefined),
  insertCagr: vi.fn(async () => undefined),
  insertColorKey: vi.fn(async () => "color key ok"),
  insertCompsStats: vi.fn(async () => "comps stats ok"),
  insertConsistentRounding: vi.fn(async () => "rounded ok"),
  insertFootballField: vi.fn(async () => "football ok"),
  insertTemplate: vi.fn(async () => "template ok"),
  insertTornado: vi.fn(async () => "tornado ok"),
  insertWaterfall: vi.fn(async () => "waterfall ok"),
  markCopySource: vi.fn(async () => "Model!A1"),
  pasteSpecial: vi.fn(async () => undefined),
  pastePreserveFormulas: vi.fn(async () => undefined),
  scaleSelection: vi.fn(async () => undefined),
  toggleIfErrorGuard: vi.fn(async () => undefined),
  undoLastAction: vi.fn(async () => "undone"),
  unpivotSelection: vi.fn(async () => "unpivoted ok"),
}));

vi.mock("./find-panel", () => ({ runFind: vi.fn(async () => "find ok") }));
vi.mock("./model-check-panel", () => ({
  runCheck: vi.fn(async () => "check ok"),
}));
vi.mock("./paint-slots", () => ({
  applyPaintSlot: vi.fn(async () => "paint apply ok"),
  capturePaintSlot: vi.fn(async () => "paint capture ok"),
}));
vi.mock("./share-panel", () => ({
  prepareShare: vi.fn(async () => "share ok"),
}));
vi.mock("./reconcile-panel", () => ({
  runReconciliation: vi.fn(async () => "reconcile ok"),
}));
vi.mock("./styles-panel", () => ({
  deleteStyles: vi.fn(async () => "styles deleted"),
  scanStyles: vi.fn(async () => "styles scanned"),
}));
vi.mock("./trace-panel", () => ({
  startTrace: vi.fn(async () => "trace ok"),
  toggleAudit: vi.fn(async () => "audit ok"),
}));
vi.mock("./workbook-tab", () => ({
  insertTocSheet: vi.fn(async () => "toc ok"),
  scanNames: vi.fn(async () => "names ok"),
}));
vi.mock("./shared", () => ({ isExcelReady: vi.fn(() => true) }));

function dataActionsInTaskpane(): string[] {
  const html = readFileSync(join(process.cwd(), "taskpane.html"), "utf8");
  const found = new Set<string>();
  for (const match of html.matchAll(/data-action="([a-zA-Z0-9_-]+)"/g)) {
    found.add(match[1]!);
  }
  return Array.from(found);
}

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(isExcelReady).mockReturnValue(true);
  Object.assign(globalThis, {
    Office: {
      context: {
        ui: {
          displayDialogAsync: vi.fn(
            (
              _url: string,
              _options: unknown,
              callback: (result: { status: string }) => void,
            ) => callback({ status: "succeeded" }),
          ),
        },
      },
      AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    },
  });
});

describe("dispatch: every taskpane data-action", () => {
  const actions = dataActionsInTaskpane();

  it("finds the full, real set of buttons (a canary for this file drifting)", () => {
    expect(actions.length).toBeGreaterThanOrEqual(60);
  });

  it.each(actions)("routes %s without throwing", async (action) => {
    await expect(dispatch(action)).resolves.toEqual(expect.any(String));
  });

  it("also routes the bespoke styles-delete confirm action", async () => {
    await expect(dispatch("styles-delete")).resolves.toEqual(
      expect.any(String),
    );
  });
});

describe("dispatch: the cycle-row-height / cycle-row- prefix shadowing bug", () => {
  it("routes cycle-row-height to the height cycle, never the row-style cycle", async () => {
    await dispatch("cycle-row-height");
    expect(applyRowHeightCycle).toHaveBeenCalledTimes(1);
    expect(applyRowStyleCycle).not.toHaveBeenCalled();
  });

  it("still routes the three real row-style kinds through the prefix branch", async () => {
    await dispatch("cycle-row-title");
    await dispatch("cycle-row-result");
    await dispatch("cycle-row-item");
    expect(applyRowStyleCycle).toHaveBeenNthCalledWith(1, "title");
    expect(applyRowStyleCycle).toHaveBeenNthCalledWith(2, "result");
    expect(applyRowStyleCycle).toHaveBeenNthCalledWith(3, "item");
    expect(applyRowHeightCycle).not.toHaveBeenCalled();
  });

  it("routes cycle-col-width to the width cycle (the same class of bug, checked)", async () => {
    await dispatch("cycle-col-width");
    expect(applyColumnWidthCycle).toHaveBeenCalledTimes(1);
  });
});

describe("dispatch: without Excel connected", () => {
  beforeEach(() => {
    vi.mocked(isExcelReady).mockReturnValue(false);
  });

  it("refuses every workbook action with one clean sentence", async () => {
    await expect(dispatch("undo")).rejects.toThrow("Excel is not connected.");
    await expect(dispatch("cycle-fill")).rejects.toThrow(
      "Excel is not connected.",
    );
    await expect(dispatch("template-dcf")).rejects.toThrow(
      "Excel is not connected.",
    );
  });

  it("still opens the shortcut card - Office chrome, not a workbook action", async () => {
    await expect(dispatch("shortcut-card")).resolves.toBe(
      "Shortcut card opened",
    );
  });

  it("falls back to a plain tab when the host has no dialog API at all", async () => {
    Object.assign(globalThis, { Office: { context: {} } });
    const opened = vi.fn();
    vi.stubGlobal("open", opened);
    await expect(dispatch("shortcut-card")).resolves.toBe(
      "Shortcut card opened",
    );
    expect(opened).toHaveBeenCalledTimes(1);
    vi.unstubAllGlobals();
  });

  it("also falls back to a plain tab when displayDialogAsync itself reports Failed", async () => {
    Object.assign(globalThis, {
      Office: {
        context: {
          ui: {
            displayDialogAsync: vi.fn(
              (
                _url: string,
                _options: unknown,
                callback: (result: { status: string }) => void,
              ) => callback({ status: "failed" }),
            ),
          },
        },
        AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      },
    });
    const opened = vi.fn();
    vi.stubGlobal("open", opened);
    await expect(dispatch("shortcut-card")).resolves.toBe(
      "Shortcut card opened",
    );
    expect(opened).toHaveBeenCalledTimes(1);
    vi.unstubAllGlobals();
  });
});

describe("dispatch: wave v2.7 slice M1", () => {
  it("routes comps-stats to the comps statistics block", async () => {
    await expect(dispatch("comps-stats")).resolves.toBe("comps stats ok");
    expect(insertCompsStats).toHaveBeenCalledTimes(1);
  });

  it("routes chart-football to the football field", async () => {
    await expect(dispatch("chart-football")).resolves.toBe("football ok");
    expect(insertFootballField).toHaveBeenCalledTimes(1);
  });

  it("routes the two pinstripe buttons to their own axis", async () => {
    await expect(dispatch("pinstripes-rows")).resolves.toBe("pinstripes ok");
    await dispatch("pinstripes-columns");
    expect(applyPinstripes).toHaveBeenNthCalledWith(1, "rows");
    expect(applyPinstripes).toHaveBeenNthCalledWith(2, "columns");
  });
});

describe("dispatch: unknown action", () => {
  it("rejects with the action name", async () => {
    await expect(dispatch("not-a-real-action")).rejects.toThrow(
      "Unknown action: not-a-real-action",
    );
  });
});
