// The football field pressed the wrong way: every selection, cap, odd value
// and host refusal from the stress brief, driven through the real pane
// dispatch. It must always answer one sentence in its own voice, and never
// leave a helper block without the chart that belongs to it.

import { beforeEach, describe, expect, it, vi } from "vitest";
import { enableStrictLoadSemantics } from "./fakehost";
import {
  AT_CAP,
  boot,
  expectSentence,
  PAST_CAP,
  planned,
  type Rig,
  seedMethods,
  sentence,
  WHOLE_SHEET,
} from "./stress.comps.support";

vi.mock("../src/pane/shared", async () =>
  (await import("./stress.comps.support")).paneShared(),
);

enableStrictLoadSemantics();

const STAGE = "Football field";
const CAP_LINE = "Football field supports up to 5,000 selected cells at once.";
const ROW_CAP_LINE = "Football field supports up to 20 rows.";
const NOT_NUMBERS =
  "Football field: the low and high columns must hold numbers";
const TOO_NARROW =
  "Football field: select at least three columns - label, low and high.";

let rig: Rig;
const run = async (): Promise<string> =>
  sentence(() => rig.dispatch("chart-football"));

/** `count` valuation methods with no header row, starting at A1. */
function seedRows(count: number): void {
  rig.helpers.seed(
    "Model!A1",
    Array.from({ length: count }, (_unused, row) => [
      `Method ${String(row)}`,
      90 + row,
      130 + row,
    ]),
  );
  rig.helpers.select(`Model!A1:C${String(count)}`);
}

beforeEach(async () => {
  rig = await boot();
});

describe("the football field on the selection a modeller leaves behind", () => {
  it("asks for three columns when one cell is selected", async () => {
    expect(await run()).toBe(TOO_NARROW);
  });

  it("asks for numbers over an empty three-column rectangle", async () => {
    rig.helpers.select("Model!C3:E6");
    expect(await run()).toBe(NOT_NUMBERS);
  });

  it("refuses a whole row, a whole column and Ctrl+A before reading", async () => {
    for (const address of ["Model!1:1", "Model!A:A", WHOLE_SHEET]) {
      rig.helpers.select(address);
      expect(await run()).toBe(CAP_LINE);
    }
    expect(rig.workbook.charts).toHaveLength(0);
    expect(rig.helpers.cellMap("Model")).toEqual({});
  });

  it("refuses a ctrl-clicked selection by name", async () => {
    seedMethods(rig);
    rig.helpers.selectAreas(["Model!A1:C4", "Model!H1:J2"]);
    expect(await run()).toBe("Football field: select a single range");
  });

  it("names the merged cell its helper block would cut", async () => {
    seedMethods(rig);
    // The block runs D1:F4; a merge straddling rows 4 and 5 is cut by it.
    rig.helpers.merge("Model!D4:D5");

    const line = await run();
    expectSentence(STAGE, line);
    expect(line).toContain("merged");
    expect(rig.workbook.charts).toHaveLength(0);
  });
});

describe("the football field at its caps", () => {
  it("lets a selection of exactly the cell cap through to the row cap", async () => {
    rig.helpers.select(AT_CAP);
    expect(await run()).toBe(ROW_CAP_LINE);
  });

  it("refuses one cell past the cell cap", async () => {
    rig.helpers.select(PAST_CAP);
    expect(await run()).toBe(CAP_LINE);
  });

  it("takes 19 and 20 rows and refuses 21", async () => {
    seedRows(19);
    expect(await run()).toBe("Football field added: 19 ranges");

    rig = await boot();
    seedRows(20);
    expect(await run()).toBe("Football field added: 20 ranges");

    rig = await boot();
    seedRows(21);
    expect(await run()).toBe(ROW_CAP_LINE);
    expect(rig.workbook.charts).toHaveLength(0);
  });

  it("takes the two-row minimum exactly", async () => {
    seedRows(1);
    expect(await run()).toBe("Football field: need at least two rows");

    rig = await boot();
    seedRows(2);
    expect(await run()).toBe("Football field added: 2 ranges");
  });

  it("refuses a selection with no room for the block beside it", async () => {
    rig.helpers.seed("Model!XEZ1", [
      ["DCF", 90, 130],
      ["Comps", 100, 120],
    ]);
    rig.helpers.select("Model!XEZ1:XFB2");
    expect(await run()).toBe(
      "Football field: no room to the right of the selection",
    );
  });
});

describe("the football field over values that are not clean numbers", () => {
  it("refuses error cells in the low or high column", async () => {
    rig.helpers.seed("Model!A1", [
      ["DCF", "#REF!", 130],
      ["Comps", 100, "#N/A"],
    ]);
    rig.helpers.select("Model!A1:C2");
    expect(await run()).toBe(NOT_NUMBERS);
    expect(rig.workbook.charts).toHaveLength(0);
  });

  it("refuses booleans and text where a valuation belongs", async () => {
    rig.helpers.seed("Model!A1", [
      ["DCF", true, 130],
      ["Comps", 100, "n/a"],
    ]);
    rig.helpers.select("Model!A1:C2");
    expect(await run()).toBe(NOT_NUMBERS);
  });

  it("draws negative and zero valuations", async () => {
    rig.helpers.seed("Model!A1", [
      ["Distressed", -50, 0],
      ["Base", 0, 120],
    ]);
    rig.helpers.select("Model!A1:C2");

    expect(await run()).toBe("Football field added: 2 ranges");
    expect(rig.helpers.value("Model!E2")).toBe(-50);
    expect(rig.helpers.value("Model!F2")).toBe(50);
  });

  it("draws a row whose low and high are the same number", async () => {
    rig.helpers.seed("Model!A1", [
      ["Point", 100, 100],
      ["Range", 90, 130],
    ]);
    rig.helpers.select("Model!A1:C2");

    expect(await run()).toBe("Football field added: 2 ranges");
    expect(rig.helpers.value("Model!F2")).toBe(0);
  });

  it("draws fifteen-digit valuations and an empty label", async () => {
    rig.helpers.seed("Model!A1", [
      [null, 999999999999999, 999999999999999],
      ["Comps", 123456789012345, 223456789012345],
    ]);
    rig.helpers.select("Model!A1:C2");

    expect(await run()).toBe("Football field added: 2 ranges");
    expect(rig.helpers.value("Model!D2")).toBe("");
    expect(rig.helpers.value("Model!F3")).toBe(100000000000000);
  });
});

describe("the football field against a host that says no", () => {
  it("names the protected sheet, spends no Undo slot and adds no chart", async () => {
    rig.helpers.select("Model!A1:C2");
    await rig.dispatch("pinstripes-rows");
    const slot = rig.smt.undoTarget();

    seedMethods(rig);
    rig.helpers.protectSheet("Model");
    expect(await run()).toBe(
      "Football field: this sheet is protected, nothing was changed",
    );
    expect(rig.smt.undoTarget()).toBe(slot);
    expect(rig.workbook.charts).toHaveLength(0);
  });

  it("spends no Undo slot when the cells beside the selection are busy", async () => {
    rig.helpers.select("Model!A1:C2");
    await rig.dispatch("pinstripes-rows");
    const slot = rig.smt.undoTarget();

    seedMethods(rig);
    rig.helpers.seed("Model!E2", [["busy"]]);
    expect(await run()).toBe(
      "Football field: cells to the right of the selection are not empty",
    );
    expect(rig.smt.undoTarget()).toBe(slot);
  });

  it("draws plain axes on a host below ExcelApi 1.7 and 1.8", async () => {
    rig.helpers.setSupported(
      (_set, version) => version !== "1.7" && version !== "1.8",
    );
    seedMethods(rig);
    expect(await run()).toBe(
      "Football field added: 3 ranges; plain axes on this build",
    );
  });

  it("says Excel placed the chart on a host below ExcelApi 1.10", async () => {
    rig.helpers.setSupported((_set, version) => version !== "1.10");
    seedMethods(rig);
    expect(await run()).toBe("Football field added: 3 ranges; Excel placed it");
  });

  it("still runs on a host with no RangeAreas (ExcelApi below 1.9)", async () => {
    rig.helpers.setSupported((_set, version) => version !== "1.9");
    seedMethods(rig);
    expect(await run()).toBe("Football field added: 3 ranges");
  });

  it("rejects rather than hangs when a read sync is refused", async () => {
    seedMethods(rig);
    rig.helpers.failNextSync();
    expect(await run()).toBe("The sync failed.");
    expect(rig.workbook.charts).toHaveLength(0);
    expect(rig.helpers.value("Model!D1")).toBe("");
  });
});

describe("the football field pressed twice", () => {
  it("answers both presses with a sentence and one helper block", async () => {
    seedMethods(rig);
    const [first, second] = await Promise.all([run(), run()]);

    expectSentence(STAGE, first);
    expectSentence(STAGE, second);
    expect(rig.helpers.value("Model!D1")).toBe("Method");
    // The block is written once over: no second block further right.
    expect(rig.helpers.value("Model!G1")).toBe("");
    // Two charts, though: both runs found the block empty and both added one.
    // Pinned here as the state of play; the skipped test at the end of this
    // file names where that is fixed.
    expect(rig.workbook.charts).toHaveLength(2);
  });

  it("puts the block back on Undo after two presses", async () => {
    seedMethods(rig);
    await run();
    await run();
    await rig.smt.undoLastAction();
    expect(rig.helpers.value("Model!D1")).toBe("");
  });
});

describe("the football field: found here, fixed elsewhere", () => {
  it.skip("adds one chart, not two, when two presses run at once", async () => {
    // Fails today with 2: both runs find the block empty and both add a
    // chart, so the sheet ends with two identical stacked bars on top of each
    // other. Not reachable from the pane - the shared guard disables every
    // [data-action] button while an action runs and this tool has no ribbon
    // command or shortcut - so it is a hazard for the day it gets one. Fix in
    // src/pane/shared.ts (not this slice): the guard should queue or drop a
    // second run, the way src/excel/link-lock.ts exclusive() serialises the
    // link flows. Taking this .skip off also flips the toHaveLength(2) in
    // "answers both presses with a sentence and one helper block" above.
    seedMethods(rig);
    await Promise.all([run(), run()]);

    expect(rig.workbook.charts).toHaveLength(1);
  });

  it.skip("draws its block over hidden rows the same as visible ones", async () => {
    // Throws today on the missing helper. Needs test/fakehost.ts (another
    // slice): rowHidden/columnHidden on FakeSheet plus helpers.hideRows.
    seedMethods(rig);
    planned(rig.helpers).hideRows("Model!3:3");

    expect(await run()).toBe("Football field added: 3 ranges");
    // The hidden method still gets its bar: the block is the selection, not
    // what the screen shows.
    expect(rig.helpers.value("Model!D3")).toBe("Trading comps");
    expect(rig.helpers.value("Model!E3")).toBe(100);
    expect(rig.helpers.value("Model!F3")).toBe(20);
  });

  it.skip("draws its block from a filtered range", async () => {
    // Throws today on the missing helper. Needs test/fakehost.ts (another
    // slice): worksheet.autoFilter with an applied range.
    seedMethods(rig);
    planned(rig.helpers).applyFilter("Model!A1:C4", "Model!3:3");

    expect(await run()).toBe("Football field added: 3 ranges");
    expect(rig.helpers.value("Model!D3")).toBe("Trading comps");
    expect(rig.workbook.charts).toHaveLength(1);
  });
});
