// Clean past the data on a sheet holding an Excel table. Before this, the
// drawings count that decides "delete" versus "clear formats only" asked for
// charts and shapes but never tables, so a table's own blank reserved rows
// (used but not filled, exactly like the ballast this tool targets) offered
// no protection: a sheet with nothing but a table would take the delete
// branch and could remove whole rows or columns that are still part of it.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["P&L", "Data"] });
  helpers = host.helpers;
  smt = await import("../../src/excel");
});

function seedStrays(rows: number, columns: number): void {
  helpers.seed("P&L!A1", [
    ["Revenue", 100, 120],
    ["Cost", -40, -50],
    ["Profit", 60, 70],
  ]);
  if (rows > 0) {
    helpers.setFill(`P&L!A4:C${String(3 + rows)}`, {
      color: "#FFEECC",
      pattern: "Solid",
    });
  }
  if (columns > 0) {
    const last = String.fromCharCode("C".charCodeAt(0) + columns);
    helpers.setFill(`P&L!D1:${last}3`, { color: "#FFEECC", pattern: "Solid" });
  }
}

describe("clean past the data with a table on the sheet", () => {
  it("clears formats only, and names tables as the reason, instead of deleting", async () => {
    seedStrays(1000, 2);
    helpers.addTable("P&L");

    expect(await smt.cleanPastData()).toBe(
      "Cleared the formats past the data on P&L; rows and columns kept because the sheet has charts, shapes or tables",
    );
    // The formatting is gone, the rows and columns themselves are still there.
    expect(helpers.fill("P&L!A4").pattern).toBe("None");
    expect(helpers.fill("P&L!D1").pattern).toBe("None");
    expect(helpers.value("P&L!C3")).toBe(70);
  });

  it("still deletes when the table sits on a different sheet", async () => {
    seedStrays(40, 0);
    helpers.addTable("Data");

    expect(await smt.cleanPastData()).toBe(
      "Removed 40 rows past the data on P&L",
    );
  });

  it("without a table at all, the same sheet still takes the delete branch", async () => {
    // Control: proves the table in the first case is what changed the
    // outcome, not some other difference in the fixture.
    seedStrays(1000, 2);

    expect(await smt.cleanPastData()).toBe(
      "Removed 1,000 rows and 2 columns past the data on P&L",
    );
  });
});
