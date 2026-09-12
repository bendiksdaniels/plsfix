// @vitest-environment jsdom
// Stress pass, slice P2: the six sheet tools and the "include very hidden"
// tick, driven the way the pane drives them - the real taskpane.html, the real
// src/pane/workbook-tab.ts, the real src/excel/workbook.ts, over the fake host.
// These tools change the sheet list, which pls,fix Undo cannot reach, so the
// line each one hands back IS the receipt: it has to be true on a one-sheet
// workbook, on a workbook whose hidden sheets are buried, on a locked
// structure and on a sheet whose name carries an apostrophe.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as WorkbookTab from "../src/pane/workbook-tab";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let tab: typeof WorkbookTab;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  const host = installFakeHost({
    sheets: ["Model", "Data", "Notes"],
    ...options,
  });
  helpers = host.helpers;
  workbook = host.workbook;
  const shared = await import("../src/pane/shared");
  shared.setExcelReady(true);
  tab = await import("../src/pane/workbook-tab");
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

function tick(on: boolean): void {
  const box = document.getElementById("sheets-very-hidden");
  (box as HTMLInputElement).checked = on;
}

/** What the sheet explorer under the buttons is showing right now. */
function explorer(): string[] {
  return Array.from(
    document.querySelectorAll("#sheet-list .sheet-row"),
    (row) => row.textContent ?? "",
  );
}

beforeEach(async () => {
  await boot();
});

describe("unhide all and the very-hidden tick", () => {
  // The line is the whole answer: a modeller who is told "1 sheet shown" and
  // still cannot find the other two has been misled, because the tick would
  // have brought them back.
  it("names the buried sheets it left even when it showed some", async () => {
    helpers.sheet("Data").visibility = "Hidden";
    helpers.sheet("Notes").visibility = "VeryHidden";
    tick(false);

    expect(await tab.unhideAllSheets()).toBe(
      '1 sheet shown. 1 more is very hidden: tick "include very hidden" for it',
    );
  });

  it("names them in the plural too, and says nothing when there are none", async () => {
    await boot({ sheets: ["A", "B", "C", "D"] });
    helpers.sheet("B").visibility = "Hidden";
    helpers.sheet("C").visibility = "VeryHidden";
    helpers.sheet("D").visibility = "VeryHidden";
    tick(false);

    expect(await tab.unhideAllSheets()).toBe(
      '1 sheet shown. 2 more are very hidden: tick "include very hidden" for them',
    );

    tick(true);
    expect(await tab.unhideAllSheets()).toBe("2 sheets shown");
  });

  it("still says so when there is nothing hidden at all", async () => {
    tick(false);
    expect(await tab.unhideAllSheets()).toBe("No hidden sheets to show");
  });

  it("points at the tick when every hidden sheet is buried", async () => {
    helpers.sheet("Data").visibility = "VeryHidden";
    tick(false);

    expect(await tab.unhideAllSheets()).toBe(
      'No hidden sheets to show. Tick "include very hidden" for the buried ones.',
    );
  });

  it("re-renders the explorer with what it just changed", async () => {
    helpers.sheet("Data").visibility = "Hidden";
    tick(true);
    await tab.unhideAllSheets();

    expect(explorer()).toHaveLength(3);
    expect(explorer().join(" ")).not.toContain("Hidden");
  });
});

describe("the tools on a workbook that leaves them nowhere to go", () => {
  it("answers every tool on a one-sheet workbook with a sentence", async () => {
    await boot({ sheets: ["Only"] });
    tick(false);

    expect(await tab.unhideAllSheets()).toBe("No hidden sheets to show");
    expect(await tab.showOnlyThisSheet()).toBe(
      "Only was already the only visible sheet",
    );
    expect(await rejects(() => tab.buryThisSheet())).toBe(
      "Excel needs one visible sheet.",
    );
    expect(await rejects(() => tab.moveThisSheet("up"))).toBe(
      "Only is already the first sheet.",
    );
    expect(await rejects(() => tab.moveThisSheet("down"))).toBe(
      "Only is already the last sheet.",
    );
    expect(await rejects(() => tab.moveThisSheet("end"))).toBe(
      "Only is already the last sheet.",
    );
  });

  it("refuses to bury the last sheet anybody can see", async () => {
    helpers.sheet("Data").visibility = "Hidden";
    helpers.sheet("Notes").visibility = "VeryHidden";

    expect(await rejects(() => tab.buryThisSheet())).toBe(
      "Excel needs one visible sheet.",
    );
    expect(helpers.sheet("Model").visibility).toBe("Visible");
  });

  it("moves past the buried sheets rather than through them", async () => {
    helpers.sheet("Data").visibility = "VeryHidden";
    workbook.activeSheetId = helpers.sheet("Notes").id;

    expect(await tab.moveThisSheet("up")).toBe("Notes is now sheet 2");
    expect(workbook.ordered().map((sheet) => sheet.name)).toEqual([
      "Model",
      "Notes",
      "Data",
    ]);
  });

  // REPORTED (P2): the receipt counts every sheet, hidden ones included, so a
  // workbook whose first sheet is hidden is told "sheet 3" about the second
  // tab it shows. Fix in src/excel/workbook.ts (not a P2 file): moveSheet
  // already loads the sheet list, so it can answer the position among the
  // VISIBLE sheets, which is the strip the modeller is counting.
  it.skip("counts the tabs a modeller can see", async () => {
    helpers.sheet("Model").visibility = "Hidden";
    workbook.activeSheetId = helpers.sheet("Data").id;

    expect(await tab.moveThisSheet("end")).toBe("Data is now sheet 2");
  });

  it("says which edge it is on rather than pretending to move", async () => {
    expect(await rejects(() => tab.moveThisSheet("up"))).toBe(
      "Model is already the first sheet.",
    );
    workbook.activeSheetId = helpers.sheet("Notes").id;
    expect(await rejects(() => tab.moveThisSheet("down"))).toBe(
      "Notes is already the last sheet.",
    );
    expect(await rejects(() => tab.moveThisSheet("end"))).toBe(
      "Notes is already the last sheet.",
    );
  });
});

describe("a sheet name a parser could trip over", () => {
  it("carries an apostrophe through every tool's line", async () => {
    await boot({ sheets: ["Bob's Model", "Data"] });
    tick(false);

    expect(await tab.showOnlyThisSheet()).toBe(
      "Only Bob's Model is visible now: 1 sheet hidden",
    );
    expect(await tab.moveThisSheet("end")).toBe("Bob's Model is now sheet 2");
    expect(await tab.unhideAllSheets()).toBe("1 sheet shown");
  });

  it("buries a sheet whose name holds a comma and an ampersand", async () => {
    await boot({ sheets: ["P&L, group", "Data"] });
    expect(await tab.buryThisSheet()).toBe(
      "P&L, group is very hidden now. Unhide all brings it back with the tick on",
    );
    expect(helpers.sheet("P&L, group").visibility).toBe("VeryHidden");
  });
});

describe("a locked sheet list", () => {
  it("answers a protected structure with the pane's sentence, per tool", async () => {
    helpers.sheet("Data").visibility = "Hidden";
    helpers.protectWorkbook();
    tick(true);
    const before = workbook.ordered().map((sheet) => sheet.name);

    expect(await rejects(() => tab.unhideAllSheets())).toBe(
      "Unhide all: this workbook's structure is protected, nothing was changed",
    );
    expect(await rejects(() => tab.showOnlyThisSheet())).toBe(
      "Show only this: this workbook's structure is protected, nothing was changed",
    );
    expect(await rejects(() => tab.buryThisSheet())).toBe(
      "Bury this sheet: this workbook's structure is protected, nothing was changed",
    );
    for (const move of ["up", "down", "end"] as const) {
      expect(await rejects(() => tab.moveThisSheet(move))).toContain(
        "this workbook's structure is protected",
      );
    }
    expect(workbook.ordered().map((sheet) => sheet.name)).toEqual(before);
    expect(helpers.sheet("Data").visibility).toBe("Hidden");
  });

  // Below ExcelApi 1.7 the structure cannot be asked about at all, so the
  // write is attempted and the host's AccessDenied is worded for the sheet
  // list rather than read as a locked cell.
  it("words the host's refusal for the sheet list below ExcelApi 1.7", async () => {
    await boot({
      sheets: ["Model", "Data"],
      isSetSupported: (set, version) =>
        !(set === "ExcelApi" && version === "1.7"),
    });
    helpers.sheet("Data").visibility = "Hidden";
    helpers.protectWorkbook();
    tick(true);

    expect(await rejects(() => tab.unhideAllSheets())).toBe(
      "Unhide all: this workbook's structure is protected, nothing was changed",
    );
    expect(await rejects(() => tab.moveThisSheet("end"))).toBe(
      "Move sheet: this workbook's structure is protected, nothing was changed",
    );
    expect(helpers.sheet("Data").visibility).toBe("Hidden");
  });

  it("leaves the cells of a protected sheet alone: these tools never touch one", async () => {
    helpers.seed("Model!A1", [["Revenue", 100]]);
    helpers.protectSheet("Model");
    workbook.activeSheetId = helpers.sheet("Data").id;
    const before = helpers.cellMap("Model");

    expect(await tab.buryThisSheet()).toContain("Data is very hidden now");
    expect(helpers.cellMap("Model")).toEqual(before);
  });
});

describe("pressed twice and pressed with a refusing host", () => {
  it("answers the second press honestly when two moves overlap", async () => {
    const both = await Promise.allSettled([
      tab.moveThisSheet("end"),
      tab.moveThisSheet("end"),
    ]);
    const said = both.map((entry) =>
      entry.status === "fulfilled"
        ? entry.value
        : (entry.reason as Error).message,
    );

    expect(said).toContain("Model is now sheet 3");
    expect(said).toContain("Model is already the last sheet.");
  });

  it("leaves the sheet list exactly as it was when the host refuses the batch", async () => {
    helpers.sheet("Data").visibility = "Hidden";
    tick(true);
    helpers.failNextSync();

    await rejects(() => tab.unhideAllSheets());
    expect(helpers.sheet("Data").visibility).toBe("Hidden");
  });

  it("spends no pls,fix undo slot: the sheet list is not cell state", async () => {
    const smt = await import("../src/excel");
    await tab.showOnlyThisSheet();
    await tab.moveThisSheet("end");
    await tab.unhideAllSheets();

    expect(smt.undoTarget()).toBeNull();
    expect(smt.lastUndoSkipped()).toBe(false);
  });
});
