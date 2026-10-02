// Comps stats pressed the wrong way: every selection, host and value a
// modeller can hand it, driven through the real pane dispatch. What it must
// always answer is one sentence in its own voice - never a raw office.js
// string, never a hang, never half a block on the sheet.

import { beforeEach, describe, expect, it, vi } from "vitest";
import { enableStrictLoadSemantics } from "./fakehost";
import {
  AT_CAP,
  boot,
  caught,
  expectSentence,
  PAST_CAP,
  REPORT_CONTEXT,
  type Rig,
  seedComps,
  sentence,
  UNDER_CAP,
  WHOLE_SHEET,
} from "./stress.comps.support";
import { describeError } from "../src/ui/report";

vi.mock("../src/pane/shared", async () =>
  (await import("./stress.comps.support")).paneShared(),
);

enableStrictLoadSemantics();

const STAGE = "Comps stats";
const CAP_LINE = "Comps stats supports up to 5,000 selected cells at once.";
const NO_NUMBERS = "Comps stats need at least one column of numbers.";
const TOO_SHORT = "Comps stats need a table with at least two data rows.";

let rig: Rig;
const run = async (): Promise<string> =>
  sentence(() => rig.dispatch("comps-stats"));

beforeEach(async () => {
  rig = await boot();
});

describe("Comps stats: the selection a modeller actually leaves behind", () => {
  it("answers a sentence on the single cell the pane opens with", async () => {
    // Nothing selected is not a state Excel has: the active cell is.
    expectSentence(STAGE, await run());
    expect(await run()).toBe(TOO_SHORT);
  });

  it("answers a sentence on an empty rectangle", async () => {
    rig.helpers.select("Model!C3:E6");
    expect(await run()).toBe(NO_NUMBERS);
  });

  it("refuses a whole row before it reads one", async () => {
    rig.helpers.select("Model!1:1");
    expect(await run()).toBe(CAP_LINE);
  });

  it("refuses a whole column before it reads one", async () => {
    rig.helpers.select("Model!A:A");
    expect(await run()).toBe(CAP_LINE);
  });

  it("refuses Ctrl+A, whose cell count comes back negative", async () => {
    rig.helpers.select(WHOLE_SHEET);
    expect(await run()).toBe(CAP_LINE);
  });

  it("refuses a ctrl-clicked selection by name", async () => {
    seedComps(rig);
    rig.helpers.selectAreas(["Model!A1:C4", "Model!E1:F2"]);
    expect(await run()).toBe("Comps stats: select a single range");
  });

  it("writes nothing at all on any refused selection", async () => {
    for (const address of ["Model!1:1", "Model!A:A", WHOLE_SHEET]) {
      rig.helpers.select(address);
      await run();
    }
    expect(rig.helpers.cellMap("Model")).toEqual({});
  });
});

describe("Comps stats at the caps", () => {
  it("lets a selection of exactly the cap through to its own rules", async () => {
    rig.helpers.select(AT_CAP);
    expect(await run()).toBe(NO_NUMBERS);
  });

  it("lets one cell under the cap through too", async () => {
    rig.helpers.select(UNDER_CAP);
    expect(await run()).toBe(NO_NUMBERS);
  });

  it("refuses one cell past the cap", async () => {
    rig.helpers.select(PAST_CAP);
    expect(await run()).toBe(CAP_LINE);
  });

  it("takes the two-data-row minimum exactly", async () => {
    rig.helpers.seed("Model!A1", [
      ["Company", "EV/EBITDA"],
      ["Alpha", 8.1],
    ]);
    rig.helpers.select("Model!A1:B2");
    expect(await run()).toBe(TOO_SHORT);

    rig.helpers.seed("Model!A3", [["Beta", 9.4]]);
    rig.helpers.select("Model!A1:B3");
    expect(await run()).toBe("Comps stats written: 1 column over 2 rows");
  });

  it("refuses a table that would run off the bottom of the sheet", async () => {
    rig.helpers.seed("Model!A1048570", [
      ["Alpha", 8.1],
      ["Beta", 9.4],
    ]);
    rig.helpers.select("Model!A1048570:B1048571");
    expect(await run()).toBe("Comps stats: no room under the selection");
  });
});

describe("Comps stats over values that are not clean numbers", () => {
  it("reads error cells as text, and refuses a column of them", async () => {
    rig.helpers.seed("Model!A1", [
      ["Company", "EV/EBITDA"],
      ["Alpha", "#REF!"],
      ["Beta", "#N/A"],
    ]);
    rig.helpers.select("Model!A1:B3");
    expect(await run()).toBe(NO_NUMBERS);
  });

  it("keeps a numeric column that holds one error cell", async () => {
    rig.helpers.seed("Model!A1", [
      ["Alpha", 8.1],
      ["Beta", "#N/A"],
      ["Gamma", 7.2],
    ]);
    rig.helpers.select("Model!A1:B3");
    expect(await run()).toBe("Comps stats written: 1 column over 3 rows");
    // The formula spans the whole column: Excel's own MIN skips the error
    // rather than the pane guessing which rows are clean.
    expect(rig.helpers.formula("Model!B5")).toBe("=MIN(B1:B3)");
  });

  it("counts zeros and negatives as numbers", async () => {
    rig.helpers.seed("Model!A1", [
      ["Alpha", 0],
      ["Beta", -4.5],
    ]);
    rig.helpers.select("Model!A1:B2");
    expect(await run()).toBe("Comps stats written: 1 column over 2 rows");
  });

  it("counts a fifteen-digit number and a date serial", async () => {
    rig.helpers.seed("Model!A1", [
      ["Alpha", 999999999999999, 45000],
      ["Beta", 123456789012345, 45300],
    ]);
    rig.helpers.setNumberFormat("Model!C1:C2", "dd/mm/yyyy");
    rig.helpers.select("Model!A1:C2");

    expect(await run()).toBe("Comps stats written: 2 columns over 2 rows");
    // The date column's statistics wear the date column's own format.
    expect(rig.helpers.numberFormat("Model!C4")).toBe("dd/mm/yyyy");
  });

  it("ignores booleans and a formula that returns an empty string", async () => {
    rig.helpers.seed("Model!A1", [
      ["Company", "Listed", "Note"],
      ["Alpha", true, { formula: '=IF(1=1,"","x")', value: "" }],
      ["Beta", false, { formula: '=IF(1=1,"","x")', value: "" }],
    ]);
    rig.helpers.select("Model!A1:C3");
    expect(await run()).toBe(NO_NUMBERS);
  });
});

describe("Comps stats against a host that says no", () => {
  it("names the protected sheet and spends no Undo slot", async () => {
    rig.helpers.select("Model!A1:C2");
    await rig.dispatch("pinstripes-rows");
    const slot = rig.smt.undoTarget();

    seedComps(rig);
    rig.helpers.protectSheet("Model");
    expect(await run()).toBe(
      "Comps stats: this sheet is protected, nothing was changed",
    );
    expect(rig.smt.undoTarget()).toBe(slot);
    expect(rig.helpers.value("Model!A6")).toBe("");
  });

  it("refuses a protected sheet even where the block would land unlocked", async () => {
    seedComps(rig);
    // The six rows are an unlocked island, so the write itself would go
    // through. The pre-check asks the sheet, not the cells - the same answer
    // Autocolor and the audit overlay have always given.
    rig.helpers.protectSheet("Model", ["Model!A6:C11"]);

    expect(await run()).toBe(
      "Comps stats: this sheet is protected, nothing was changed",
    );
    expect(rig.helpers.value("Model!A6")).toBe("");
  });

  it("names the merged cell its block would cut in half", async () => {
    seedComps(rig);
    // The six rows start at row 6; a merge straddling rows 5 and 6 is cut.
    rig.helpers.merge("Model!A5:A6");

    const line = await run();
    expectSentence(STAGE, line);
    expect(line).toContain("merged");
  });

  it("still runs on a host with no RangeAreas (ExcelApi below 1.9)", async () => {
    rig.helpers.setSupported((_set, version) => version !== "1.9");
    seedComps(rig);
    expect(await run()).toBe("Comps stats written: 1 column over 3 rows");
  });

  it("writes nothing when the host cannot be asked about protection", async () => {
    // Below ExcelApi 1.2 sheetProtected answers false, so the refusal has to
    // arrive from the write itself - still as this tool's own sentence.
    rig.helpers.setSupported((_set, version) => version !== "1.2");
    seedComps(rig);
    rig.helpers.protectSheet("Model");
    expect(await run()).toBe(
      "Comps stats: this sheet is protected, nothing was changed",
    );
    expect(rig.helpers.value("Model!A6")).toBe("");
  });

  it("rejects rather than hangs when a read sync is refused", async () => {
    seedComps(rig);
    rig.helpers.failNextSync();
    // A host that drops a read batch has no staged wording of its own yet:
    // see the skipped test at the end of this file.
    expect(await run()).toBe("The sync failed.");
    expect(rig.helpers.value("Model!A6")).toBe("");
  });

  it("writes nothing at all when the host refuses the block", async () => {
    seedComps(rig);
    rig.helpers.merge("Model!A5:A6");
    expectSentence(STAGE, await run());

    expect(rig.helpers.value("Model!A6")).toBe("");
    expect(rig.helpers.value("Model!B6")).toBe("");
    expect(rig.helpers.value("Model!B11")).toBe("");
  });
});

describe("Comps stats pressed twice", () => {
  it("leaves one whole block when two presses run at once", async () => {
    seedComps(rig);
    const [first, second] = await Promise.all([run(), run()]);

    expectSentence(STAGE, first);
    expectSentence(STAGE, second);
    // Both presses read the same empty block, so both write the same six
    // rows: the second is idempotent, never a second block or half of one.
    expect(first).toBe(second);
    // One of the two writes the block; whichever ran second either wrote the
    // same six rows again or was refused for finding them there.
    expect(rig.helpers.value("Model!A6")).toBe("Min");
    expect(rig.helpers.formula("Model!B11")).toBe("=MAX(B2:B4)");
    expect(rig.helpers.value("Model!A12")).toBe("");
  });

  it("clears the block once, however many presses landed", async () => {
    seedComps(rig);
    await run();
    await run();
    await rig.smt.undoLastAction();
    expect(rig.helpers.value("Model!A6")).toBe("");
  });
});

describe("Comps stats: found here, fixed elsewhere", () => {
  it("spends no Undo slot when the host refuses the write itself", async () => {
    // Fails today at the last line with "Model!A6:C11": the write comes back
    // InvalidOperation, but captureUndo has already pushed the block onto the
    // five-deep stack, dropping the modeller's oldest entry - and that new
    // entry can never restore while the merge stands. Every writing flow in
    // src/excel does this, not just the comps tools. Fix in src/excel/undo.ts
    // (not this slice): a discardUndo() the flow calls when its write comes
    // back refused, or a captureUndo that only commits once the batch lands.
    rig.helpers.select("Model!A1:C2");
    await rig.dispatch("pinstripes-rows");
    const slot = rig.smt.undoTarget();

    seedComps(rig);
    rig.helpers.merge("Model!A5:A6");
    expectSentence(STAGE, await run());

    expect(rig.smt.undoTarget()).toBe(slot);
  });

  it("writes its block over hidden rows the same as visible ones", async () => {
    seedComps(rig);
    rig.helpers.hideRows("Model!3:3");

    expect(await run()).toBe("Comps stats written: 1 column over 3 rows");
    // The span covers the hidden row: Excel's own MIN counts it, so the block
    // must not shrink to what the screen shows.
    expect(rig.helpers.formula("Model!B6")).toBe("=MIN(B2:B4)");
    expect(rig.helpers.value("Model!A6")).toBe("Min");
  });

  it("writes its block under a filtered range", async () => {
    seedComps(rig);
    rig.helpers.applyFilter("Model!A1:C4", "Model!3:3");

    expect(await run()).toBe("Comps stats written: 1 column over 3 rows");
    // Under the table's last row, not under the last visible one.
    expect(rig.helpers.value("Model!A6")).toBe("Min");
    expect(rig.helpers.formula("Model!B6")).toBe("=MIN(B2:B4)");
  });

  it("stages a host error a read sync drops", async () => {
    // "The sync failed." is a bare host string. Every read sync in src/excel
    // is bare by design - only write syncs stage their own refusal, through
    // syncWrite/paintSync - so this is a pane-wide policy, not a comps one:
    // describeError (src/ui/report.ts) now prefixes the running action's
    // label, the one pure half of the pair these suites can reach - they mock
    // src/pane/shared.ts's guard away.
    seedComps(rig);
    rig.helpers.failNextSync();
    const error = await caught(() => rig.dispatch("comps-stats"));

    const { message } = describeError(error, REPORT_CONTEXT, "comps-stats");
    expectSentence(STAGE, message);
  });
});
