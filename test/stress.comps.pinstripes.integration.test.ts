// Pinstripes pressed the wrong way, both axes: every selection, cap, host
// refusal and repeat press from the stress brief, driven through the real pane
// dispatch. A band is a reading aid, so the answer is always a sentence - even
// where the sheet refuses the paint - and never a raw office.js string.

import { beforeEach, describe, expect, it, vi } from "vitest";
import { enableStrictLoadSemantics } from "./fakehost";
import {
  AT_CAP,
  boot,
  expectSentence,
  PAST_CAP,
  type Rig,
  seedGrid,
  sentence,
  UNDER_CAP,
  WHOLE_SHEET,
} from "./stress.comps.support";
import { DEFAULT_SETTINGS, tint } from "../src/settings";

vi.mock("../src/pane/shared", async () =>
  (await import("./stress.comps.support")).paneShared(),
);

enableStrictLoadSemantics();

const STAGE = "Pinstripes";
const BAND = tint(DEFAULT_SETTINGS.primary, 0.9);
const CAP_LINE = "Pinstripes supports up to 5,000 selected cells at once.";
const PROTECTED = "Pinstripes: this sheet is protected, nothing was changed";

let rig: Rig;
const rows = async (): Promise<string> =>
  sentence(() => rig.dispatch("pinstripes-rows"));
const columns = async (): Promise<string> =>
  sentence(() => rig.dispatch("pinstripes-columns"));

beforeEach(async () => {
  rig = await boot();
});

describe("Pinstripes on the selection a modeller leaves behind", () => {
  it("asks for two lines when one cell is selected", async () => {
    expect(await rows()).toBe(
      "Pinstripes need at least two rows in the selection.",
    );
    expect(await columns()).toBe(
      "Pinstripes need at least two columns in the selection.",
    );
  });

  it("bands an empty rectangle, since a band is formatting", async () => {
    rig.helpers.select("Model!C3:E6");
    expect(await rows()).toBe("Pinstripes: 2 rows banded");
    expect(rig.helpers.fill("Model!C4").color).toBe(BAND);
  });

  it("refuses a whole row, a whole column and Ctrl+A before reading a fill", async () => {
    for (const address of ["Model!1:1", "Model!A:A", WHOLE_SHEET]) {
      rig.helpers.select(address);
      expect(await rows()).toBe(CAP_LINE);
      expect(await columns()).toBe(CAP_LINE);
    }
    expect(rig.helpers.cellMap("Model")).toEqual({});
  });

  it("refuses a ctrl-clicked selection by name, on both axes", async () => {
    seedGrid(rig);
    rig.helpers.selectAreas(["Model!A1:C5", "Model!E1:F2"]);
    expect(await rows()).toBe("Pinstripes: select a single range");
    expect(await columns()).toBe("Pinstripes: select a single range");
  });

  it("bands across the columns of a single row", async () => {
    rig.helpers.seed("Model!A1", [[1, 2, 3, 4]]);
    rig.helpers.select("Model!A1:D1");
    expect(await columns()).toBe("Pinstripes: 2 columns banded");
    expect(rig.helpers.fill("Model!B1").color).toBe(BAND);
    expect(rig.helpers.fill("Model!D1").color).toBe(BAND);
  });

  it("bands over a merged block a banded line cuts in half", async () => {
    rig.helpers.seed("Model!A1", [[1, 2]]);
    rig.helpers.merge("Model!A2:A3");
    rig.helpers.seed("Model!A4", [[3, 4]]);
    rig.helpers.select("Model!A1:B4");

    // Row 2 is banded and row 3 is not, so the paint covers half the merge -
    // a fill, not a value write, which Excel takes.
    expect(await rows()).toBe("Pinstripes: 2 rows banded");
    expect(rig.helpers.fill("Model!A2").color).toBe(BAND);
    expect(rig.helpers.fill("Model!A3").pattern).toBe("None");
  });
});

describe("Pinstripes at the cap", () => {
  it("bands a selection of exactly the cap", async () => {
    rig.helpers.select(AT_CAP);
    expect(await rows()).toBe("Pinstripes: 500 rows banded");
  });

  it("bands one cell under the cap", async () => {
    rig.helpers.select(UNDER_CAP);
    expectSentence(STAGE, await rows());
  });

  it("refuses one cell past the cap", async () => {
    rig.helpers.select(PAST_CAP);
    expect(await rows()).toBe(CAP_LINE);
  });

  it("bands the two-line minimum exactly, in the singular", async () => {
    rig.helpers.seed("Model!A1", [
      [1, 2],
      [3, 4],
    ]);
    rig.helpers.select("Model!A1:B2");
    expect(await rows()).toBe("Pinstripes: 1 row banded");
  });
});

describe("Pinstripes over values it never reads", () => {
  it("bands error cells, dates, booleans and blanks alike", async () => {
    rig.helpers.seed("Model!A1", [
      ["#REF!", 45000, true],
      ["#N/A", null, false],
      ["", -0.5, 999999999999999],
      ["x", 0, ""],
    ]);
    rig.helpers.select("Model!A1:C4");

    expect(await rows()).toBe("Pinstripes: 2 rows banded");
    expect(rig.helpers.fill("Model!A2").color).toBe(BAND);
    expect(rig.helpers.fill("Model!C4").color).toBe(BAND);
  });
});

describe("Pinstripes against a host that says no", () => {
  it("says the sheet is protected, paints nothing and keeps the Undo slot", async () => {
    rig.helpers.select("Model!A1:C2");
    await rows();
    const slot = rig.smt.undoTarget();

    rig.helpers.select("Data!A1:C5");
    rig.helpers.protectSheet("Data");
    expect(await rows()).toBe(PROTECTED);
    expect(rig.smt.undoTarget()).toBe(slot);
    expect(rig.helpers.fill("Data!A2").pattern).toBe("None");
  });

  it("says the same when the host is too old to be asked (below 1.2)", async () => {
    rig.helpers.setSupported((_set, version) => version !== "1.2");
    seedGrid(rig);
    rig.helpers.protectSheet("Model");
    expect(await rows()).toBe(PROTECTED);
    expect(rig.helpers.fill("Model!A2").pattern).toBe("None");
  });

  it("still runs on a host with no RangeAreas (ExcelApi below 1.9)", async () => {
    rig.helpers.setSupported((_set, version) => version !== "1.9");
    seedGrid(rig);
    expect(await rows()).toBe("Pinstripes: 2 rows banded");
  });

  it("rejects rather than hangs when a read sync is refused", async () => {
    seedGrid(rig);
    rig.helpers.failNextSync();
    expect(await rows()).toBe("The sync failed.");
    expect(rig.helpers.fill("Model!A2").pattern).toBe("None");
  });

  it("paints nothing when the host refuses the fill read", async () => {
    seedGrid(rig);
    rig.helpers.failNextCellProperties();
    // The host's own string, with no tool name on it: the same gap as the
    // refused read sync above, reported for src/pane/shared.ts.
    expect(await rows()).toBe("The cell formats failed.");
    expect(rig.helpers.fill("Model!A2").pattern).toBe("None");
    expect(rig.smt.undoTarget()).toBeNull();
  });
});

describe("Pinstripes pressed twice", () => {
  it("bands and then clears, whichever axis", async () => {
    seedGrid(rig);
    expect(await rows()).toBe("Pinstripes: 2 rows banded");
    expect(await rows()).toBe("Pinstripes: 2 rows cleared");
    expect(await columns()).toBe("Pinstripes: 1 column banded");
    expect(await columns()).toBe("Pinstripes: 1 column cleared");
  });

  it("bands both ways when the axes are crossed, and Undo takes it back", async () => {
    seedGrid(rig);
    await rows();
    expect(await columns()).toBe("Pinstripes: 1 column banded");
    expect(rig.helpers.fill("Model!B1").color).toBe(BAND);

    await rig.smt.undoLastAction();
    await rig.smt.undoLastAction();
    expect(rig.helpers.fill("Model!A2").pattern).toBe("None");
    expect(rig.helpers.fill("Model!B1").pattern).toBe("None");
  });

  it("lands on the same bands when two presses run at once", async () => {
    seedGrid(rig);
    const [first, second] = await Promise.all([rows(), rows()]);

    expectSentence(STAGE, first);
    expectSentence(STAGE, second);
    // Both read the fills before either painted, so both band: the second is
    // idempotent rather than a clear of the first.
    expect(first).toBe("Pinstripes: 2 rows banded");
    expect(second).toBe("Pinstripes: 2 rows banded");
    expect(rig.helpers.fill("Model!A2").color).toBe(BAND);
    expect(rig.helpers.fill("Model!A1").pattern).toBe("None");
  });

  it("puts the modeller's own fills back after both presses", async () => {
    seedGrid(rig);
    rig.helpers.setFill("Model!A2", { color: "#FFFF00", pattern: "Solid" });
    await Promise.all([rows(), rows()]);

    await rig.smt.undoLastAction();
    await rig.smt.undoLastAction();
    expect(rig.helpers.fill("Model!A2").color).toBe("#FFFF00");
  });
});

describe.skip("Pinstripes: found here, fixed elsewhere", () => {
  it("stages a host error a read refuses", () => {
    // Proven by the two tests above: "The sync failed." and "The cell formats
    // failed." reach the toast verbatim, with no tool name on them, because
    // the guard hands error.message straight through. Every read sync in
    // src/excel is the same. Fix in src/pane/shared.ts / src/ui/report.ts
    // (not this slice): describeError should prefix the running action's
    // label when the message carries no stage of its own.
  });

  it("bands a hidden row the same as a visible one", () => {
    // Needs test/fakehost.ts (another slice): no row or column visibility, so
    // a hidden line inside the selection cannot be modelled. Real Excel bands
    // it, and unhiding shows the band, which is what a modeller expects.
    // Fix: rowHidden/columnHidden on FakeSheet plus helpers.hideRows(address).
  });

  it("bands the rows a filter left showing", () => {
    // Needs test/fakehost.ts (another slice): no AutoFilter surface at all.
    // Worth a decision as well as a fake: Excel's own banded-table style
    // follows the filter, our every-second-row does not.
  });
});
