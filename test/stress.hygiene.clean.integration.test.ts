// Stress pass, slice P2: Clean past the data, the one hygiene tool that
// DELETES. It runs outside pls,fix Undo and Office.js writes never reach
// Excel's own undo stack, so a wrong delete is gone for good - which makes the
// interesting abuse the shapes of used range a saved workbook really carries:
// one cell, formats and no values, a merge across the edge, a million
// formatted rows, a chart nobody can see, and the button pressed twice.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"], ...options });
  helpers = host.helpers;
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

const KEPT_FOR_DRAWINGS =
  "rows and columns kept because the sheet has charts or shapes";

/** Two rows of model, then whatever ballast the scenario wants around it. */
function seedModel(): void {
  helpers.seed("Model!A1", [
    ["Revenue", 100],
    ["Cost", -40],
  ]);
}

function ballastRows(through: number): void {
  helpers.setFill(`Model!A3:B${String(through)}`, {
    color: "#FFEECC",
    pattern: "Solid",
  });
}

beforeEach(async () => {
  await boot();
});

describe("used ranges a saved workbook really carries", () => {
  it("says so when the used range is a single cell", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    expect(await smt.cleanPastData()).toBe("Nothing past the data on Model.");
    expect(helpers.value("Model!A1")).toBe("Revenue");
  });

  it("leaves a sheet whose used range is formats and no values", async () => {
    helpers.setFill("Model!A1:C50", { color: "#FFEECC", pattern: "Solid" });
    expect(await smt.cleanPastData()).toBe("Nothing on this sheet to clean.");
    expect(helpers.fill("Model!A1").pattern).toBe("Solid");
    expect(helpers.fill("Model!C50").pattern).toBe("Solid");
  });

  it("keeps the formatting above and left of data that does not start at A1", async () => {
    helpers.seed("Model!C5", [["Revenue", 100]]);
    helpers.setFill("Model!A1:H40", { color: "#FFEECC", pattern: "Solid" });

    expect(await smt.cleanPastData()).toBe(
      "Removed 35 rows and 4 columns past the data on Model",
    );
    expect(helpers.value("Model!C5")).toBe("Revenue");
    expect(helpers.fill("Model!A1").pattern).toBe("Solid");
  });

  // Excel answers cellCount -1 past 2^31-1 cells, and a whole sheet of
  // formatting is exactly that. The tool never counts cells: it deletes bands.
  it("clears a million formatted rows in one pass", async () => {
    helpers.seed("Model!A1", [["Revenue", 100]]);
    helpers.setFill("Model!A2:A1048576", {
      color: "#FFEECC",
      pattern: "Solid",
    });

    expect(await smt.cleanPastData()).toBe(
      "Removed 1,048,575 rows past the data on Model",
    );
    expect(helpers.value("Model!B1")).toBe(100);
  }, 30_000);

  it("counts only the columns when the ballast is beside the data", async () => {
    seedModel();
    helpers.setFill("Model!C1:F2", { color: "#FFEECC", pattern: "Solid" });

    expect(await smt.cleanPastData()).toBe(
      "Removed 4 columns past the data on Model",
    );
    expect(helpers.value("Model!B2")).toBe(-40);
  });

  it("says nothing to clean on a sheet with nothing on it", async () => {
    expect(await smt.cleanPastData()).toBe("Nothing on this sheet to clean.");
  });

  // A formula that evaluates to "" is still something a modeller typed: the
  // cell counts as data and the band under it is what goes.
  it("counts a formula returning empty text as data", async () => {
    helpers.seed("Model!A1", [[{ formula: '=IF(1=1,"","x")', value: "" }]]);
    helpers.setFill("Model!A2:A9", { color: "#FFEECC", pattern: "Solid" });

    expect(await smt.cleanPastData()).toBe(
      "Removed 8 rows past the data on Model",
    );
    expect(helpers.formula("Model!A1")).toBe('=IF(1=1,"","x")');
  });
});

describe("things a delete could destroy", () => {
  // A merge whose top-left sits in the data and whose tail runs past it: the
  // value lives in the top-left, so the delete may cut the tail but never the
  // number. (What Excel leaves of the merge itself is a launch check.)
  it("keeps the value of a merge that straddles the data edge", async () => {
    helpers.seed("Model!A1", [
      ["Revenue", 100],
      ["Note", 0],
    ]);
    helpers.merge("Model!A2:A4");
    ballastRows(9);

    expect(await smt.cleanPastData()).toBe(
      "Removed 7 rows past the data on Model",
    );
    expect(helpers.value("Model!A2")).toBe("Note");
    expect(helpers.value("Model!B1")).toBe(100);
  });

  it("takes a merge that sits entirely past the data with the rows", async () => {
    seedModel();
    ballastRows(9);
    helpers.merge("Model!A6:B6");

    expect(await smt.cleanPastData()).toBe(
      "Removed 7 rows past the data on Model",
    );
    expect(helpers.sheet("Model").merges).toEqual([]);
    expect(helpers.value("Model!A2")).toBe("Cost");
  });

  it("only strips the formats while a chart could be sitting on the rows", async () => {
    seedModel();
    ballastRows(9);
    helpers.merge("Model!A6:B6");
    helpers.addChart("Model");

    expect(await smt.cleanPastData()).toBe(
      `Cleared the formats past the data on Model; ${KEPT_FOR_DRAWINGS}`,
    );
    // The witness that nothing was deleted: the merge is still standing.
    expect(helpers.sheet("Model").merges).toHaveLength(1);
    expect(helpers.fill("Model!A3").pattern).toBe("None");
  });

  it("never strips the data's own formatting on the safe branch", async () => {
    seedModel();
    helpers.setFill("Model!A1:B2", { color: "#112233", pattern: "Solid" });
    ballastRows(9);
    helpers.addChart("Model");

    await smt.cleanPastData();
    expect(helpers.fill("Model!A1").pattern).toBe("Solid");
    expect(helpers.fill("Model!A1").color).toBe("#112233");
  });

  // A host that cannot be asked how many charts or shapes a sheet holds takes
  // the branch that destroys nothing: fail closed, both ways round.
  it("fails closed on a host that cannot count charts or shapes", async () => {
    for (const apiSet of ["1.4", "1.9"]) {
      await boot({
        isSetSupported: (set, version) =>
          !(set === "ExcelApi" && version === apiSet),
      });
      seedModel();
      ballastRows(9);
      helpers.merge("Model!A6:B6");

      expect(await smt.cleanPastData()).toBe(
        "Cleared the formats past the data on Model; rows and columns kept because this Excel cannot count the sheet's charts or shapes",
      );
      expect(helpers.sheet("Model").merges).toHaveLength(1);
    }
  });

  it("deletes once the chart is gone and the host can say so", async () => {
    seedModel();
    ballastRows(9);
    helpers.addChart("Model");
    await smt.cleanPastData();

    await boot();
    seedModel();
    ballastRows(9);
    expect(await smt.cleanPastData()).toBe(
      "Removed 7 rows past the data on Model",
    );
  });

  it("cleans the sheet the workbook is looking at and no other", async () => {
    seedModel();
    ballastRows(9);
    helpers.setFill("Data!A1:B40", { color: "#FFEECC", pattern: "Solid" });

    await smt.cleanPastData();
    expect(helpers.fill("Data!A40").pattern).toBe("Solid");
  });
});

describe("the wrong time to press it", () => {
  it("answers a protected sheet with the pane's sentence and deletes nothing", async () => {
    seedModel();
    ballastRows(9);
    helpers.protectSheet("Model");
    const before = helpers.cellMap("Model");

    expect(await rejects(() => smt.cleanPastData())).toBe(
      "Clean past the data: this sheet is protected, nothing was changed",
    );
    expect(helpers.cellMap("Model")).toEqual(before);
  });

  it("leaves the sheet whole when the host refuses the batch", async () => {
    seedModel();
    ballastRows(9);
    helpers.failNextSync();

    await rejects(() => smt.cleanPastData());
    expect(helpers.value("Model!A2")).toBe("Cost");
    expect(helpers.value("Model!B2")).toBe(-40);
  });

  it("says there is nothing left on the second and third press", async () => {
    seedModel();
    ballastRows(9);

    expect(await smt.cleanPastData()).toBe(
      "Removed 7 rows past the data on Model",
    );
    expect(await smt.cleanPastData()).toBe("Nothing past the data on Model.");
    expect(await smt.cleanPastData()).toBe("Nothing past the data on Model.");
  });

  // Two presses that overlap both scan the same used range and both delete the
  // same band; the second one finds it empty. The data survives either way -
  // the surplus is by definition everything BELOW and RIGHT of the values.
  it("keeps the data when two presses overlap", async () => {
    seedModel();
    helpers.setFill("Model!C1:F9", { color: "#FFEECC", pattern: "Solid" });

    const both = await Promise.allSettled([
      smt.cleanPastData(),
      smt.cleanPastData(),
    ]);
    expect(both.map((entry) => entry.status)).toEqual([
      "fulfilled",
      "fulfilled",
    ]);
    expect(helpers.value("Model!A1")).toBe("Revenue");
    expect(helpers.value("Model!B2")).toBe(-40);
  });

  it("ignores the selection, whatever shape it is in", async () => {
    seedModel();
    ballastRows(9);
    helpers.merge("Model!A1:B1");
    helpers.selectAreas(["Model!A1:B1", "Model!D1:D1048576"]);

    expect(await smt.cleanPastData()).toBe(
      "Removed 7 rows past the data on Model",
    );
  });

  it("spends no pls,fix undo slot and never claims one was skipped", async () => {
    seedModel();
    ballastRows(9);
    await smt.cleanPastData();

    expect(smt.undoTarget()).toBeNull();
    expect(smt.lastUndoSkipped()).toBe(false);
  });
});
