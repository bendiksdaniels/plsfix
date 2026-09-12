// The Workbook tab's sheet tools against the fake host: unhide all, show only
// this, bury this and the three moves. These act on the workbook's sheet list,
// never on the selection, and they run outside pls,fix Undo - sheet state is
// not cell state and getCellProperties carries none of it.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(sheets = ["Model", "Data", "Notes"]): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

function visibility(name: string): string {
  return helpers.sheet(name).visibility;
}

function order(): string[] {
  return workbook.ordered().map((sheet) => sheet.name);
}

beforeEach(async () => {
  await boot();
});

describe("unhide all", () => {
  it("shows every hidden sheet and counts them", async () => {
    helpers.sheet("Data").visibility = "Hidden";
    helpers.sheet("Notes").visibility = "Hidden";

    expect(await smt.setSheetsVisibility(false)).toBe(2);
    expect(visibility("Data")).toBe("Visible");
    expect(visibility("Notes")).toBe("Visible");
  });

  it("leaves a very hidden sheet alone unless it is asked for it", async () => {
    helpers.sheet("Data").visibility = "Hidden";
    helpers.sheet("Notes").visibility = "VeryHidden";

    expect(await smt.setSheetsVisibility(false)).toBe(1);
    expect(visibility("Notes")).toBe("VeryHidden");

    expect(await smt.setSheetsVisibility(true)).toBe(1);
    expect(visibility("Notes")).toBe("Visible");
  });

  it("counts nothing when every sheet is already showing", async () => {
    expect(await smt.setSheetsVisibility(true)).toBe(0);
  });
});

describe("show only this", () => {
  it("hides every other visible sheet and names the one that is left", async () => {
    expect(await smt.showOnlySheet()).toEqual({ name: "Model", hidden: 2 });
    expect(visibility("Model")).toBe("Visible");
    expect(visibility("Data")).toBe("Hidden");
    expect(visibility("Notes")).toBe("Hidden");
  });

  it("shows the active sheet even when it was hidden itself", async () => {
    helpers.sheet("Data").visibility = "Hidden";
    workbook.activeSheetId = helpers.sheet("Data").id;

    expect(await smt.showOnlySheet()).toEqual({ name: "Data", hidden: 2 });
    expect(visibility("Data")).toBe("Visible");
  });

  it("does not disturb a very hidden sheet", async () => {
    helpers.sheet("Notes").visibility = "VeryHidden";

    expect(await smt.showOnlySheet()).toEqual({ name: "Model", hidden: 1 });
    expect(visibility("Notes")).toBe("VeryHidden");
  });
});

describe("bury this", () => {
  it("makes the active sheet very hidden", async () => {
    workbook.activeSheetId = helpers.sheet("Data").id;

    expect(await smt.burySheet()).toBe("Data");
    expect(visibility("Data")).toBe("VeryHidden");
  });

  it("refuses to bury the last visible sheet", async () => {
    helpers.sheet("Data").visibility = "Hidden";
    helpers.sheet("Notes").visibility = "Hidden";

    expect(await rejects(() => smt.burySheet())).toBe(
      "Excel needs one visible sheet.",
    );
    expect(visibility("Model")).toBe("Visible");
  });

  it("refuses on a one-sheet workbook", async () => {
    await boot(["Model"]);
    expect(await rejects(() => smt.burySheet())).toBe(
      "Excel needs one visible sheet.",
    );
  });
});

describe("moving a sheet", () => {
  it("moves the active sheet up, down and to the end", async () => {
    workbook.activeSheetId = helpers.sheet("Data").id;

    expect(await smt.moveSheet("up")).toEqual({ name: "Data", position: 0 });
    expect(order()).toEqual(["Data", "Model", "Notes"]);

    expect(await smt.moveSheet("down")).toEqual({ name: "Data", position: 1 });
    expect(order()).toEqual(["Model", "Data", "Notes"]);

    expect(await smt.moveSheet("end")).toEqual({ name: "Data", position: 2 });
    expect(order()).toEqual(["Model", "Notes", "Data"]);
  });

  it("says so rather than pretending when there is nowhere to go", async () => {
    expect(await rejects(() => smt.moveSheet("up"))).toBe(
      "Model is already the first sheet.",
    );

    workbook.activeSheetId = helpers.sheet("Notes").id;
    expect(await rejects(() => smt.moveSheet("down"))).toBe(
      "Notes is already the last sheet.",
    );
    expect(await rejects(() => smt.moveSheet("end"))).toBe(
      "Notes is already the last sheet.",
    );
    expect(order()).toEqual(["Model", "Data", "Notes"]);
  });
});

describe("the universality rows", () => {
  // These tools read the sheet list, never the selection: a ctrl-clicked
  // selection, a merged block and a whole-column click all reach the same
  // answer, and none of them is refused for its size.
  it("ignores the selection, whatever shape it is in", async () => {
    helpers.merge("Model!A1:B1");
    helpers.selectAreas(["Model!A1:B1", "Model!D1:D5000"]);

    expect(await smt.showOnlySheet()).toEqual({ name: "Model", hidden: 2 });
  });

  it("hides and moves sheets on a protected sheet: cells are not touched", async () => {
    helpers.protectSheet("Model");
    workbook.activeSheetId = helpers.sheet("Data").id;

    expect(await smt.burySheet()).toBe("Data");
    expect(await smt.setSheetsVisibility(true)).toBe(1);
  });

  it("answers a protected workbook structure with the pane's sentence", async () => {
    helpers.protectWorkbook();

    expect(await rejects(() => smt.setSheetsVisibility(true))).toBe(
      "Unhide all: this workbook's structure is protected, nothing was changed",
    );
    expect(await rejects(() => smt.showOnlySheet())).toBe(
      "Show only this: this workbook's structure is protected, nothing was changed",
    );
    expect(await rejects(() => smt.burySheet())).toBe(
      "Bury this sheet: this workbook's structure is protected, nothing was changed",
    );
    expect(await rejects(() => smt.moveSheet("end"))).toBe(
      "Move sheet: this workbook's structure is protected, nothing was changed",
    );
    expect(order()).toEqual(["Model", "Data", "Notes"]);
  });

  it("leaves the workbook exactly as it was when the host refuses", async () => {
    helpers.sheet("Data").visibility = "Hidden";
    helpers.protectWorkbook();

    await rejects(() => smt.setSheetsVisibility(true));
    expect(visibility("Data")).toBe("Hidden");
  });

  it("spends no pls,fix undo slot: sheet state is not cell state", async () => {
    await smt.showOnlySheet();
    await smt.moveSheet("end");
    expect(smt.undoTarget()).toBeNull();
  });
});
