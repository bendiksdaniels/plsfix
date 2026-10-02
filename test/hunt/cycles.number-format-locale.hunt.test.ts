// Pass-1 hunt: canonicalNumberFormat against a bracket Excel rewrites with no
// currency symbol in it at all - a bare locale tag, [$-409], the form Excel
// uses on read-back for any format whose display depends on the regional
// settings (a month name, not only a currency): the "never exact-match state
// Excel gives back" lesson (2026-08-27) covered the tagged-symbol case
// ([$€-x-euro2]) but the regex still required at least one character before
// the dash, so a bare tag survived canonicalisation untouched.
import { describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";
import {
  buildNumberCycles,
  canonicalNumberFormat,
  nextInCycle,
} from "../../src/cycles";
import { DEFAULT_SETTINGS } from "../../src/settings";

enableStrictLoadSemantics();

describe("canonicalNumberFormat against a bare locale tag", () => {
  it("strips a locale-only bracket the same way it strips a tagged symbol", () => {
    expect(canonicalNumberFormat("[$-409]dd.mm.yyyy")).toBe("dd.mm.yyyy");
    expect(canonicalNumberFormat("[$-409]mmm-yy;@")).toBe("mmm-yy;@");
  });

  it("still strips the tagged-symbol form the lesson already covered", () => {
    expect(canonicalNumberFormat("[$€-x-euro2] #,##0")).toBe("€ #,##0");
  });
});

describe("the date cycle stepping past a bare locale tag on read-back", () => {
  const dateCycle = buildNumberCycles(DEFAULT_SETTINGS).date;

  it("steps to the next date format instead of restarting at the first", () => {
    // "mmm-yy" read back tagged, as Excel does for a format whose month
    // abbreviation depends on the locale - never our own literal.
    expect(nextInCycle("[$-409]mmm-yy", dateCycle)).toBe(dateCycle[2]);
  });
});

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({
    sheets: ["Model", "Data"],
    rewriteDateLocaleTag: true,
  });
  helpers = host.helpers;
  smt = await import("../../src/excel");
}

// The three date formats a modeller actually gets, three presses in a row,
// against a host that tags every one of them with a bare locale code -
// never sticking on the first rung, never restarting - then Undo back to
// the exact original, byte for byte.
describe("the date format cycle over a host that tags every read-back", () => {
  it("steps forward on all three presses and fully undoes", async () => {
    await boot();
    helpers.setNumberFormat("Model!A1", "General");
    helpers.select("Model!A1");
    const before = helpers.cellMap("Model");

    await smt.applyNumberCycle("date");
    const first = helpers.numberFormat("Model!A1");
    expect(first).toContain("[$-409]");
    expect(first).toContain("dd.mm.yyyy");

    await smt.applyNumberCycle("date");
    const second = helpers.numberFormat("Model!A1");
    expect(second).toContain("[$-409]");
    expect(second).toContain("mmm-yy");
    expect(second).not.toBe(first);

    await smt.applyNumberCycle("date");
    const third = helpers.numberFormat("Model!A1");
    expect(third).toContain("[$-409]");
    expect(third).toContain("yyyy");
    expect(third).not.toBe(second);

    await smt.undoLastAction();
    await smt.undoLastAction();
    await smt.undoLastAction();
    expect(helpers.cellMap("Model")).toEqual(before);
  });
});
