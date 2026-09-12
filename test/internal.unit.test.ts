// Direct proof of the pure-enough helpers inside src/excel/internal.ts: the
// cell caps (overCap, pickScannableSheets), the number-format and fill-key
// round trips, the host-capability probe and the broken-name filter behind
// Super Find. Host-dependent exports (withinCap, styleChartShell,
// syncTolerating, ...) already have coverage in the *.integration.test.ts
// suites and stay there - this file never opens a real (or fake) Excel.run.
import { afterAll, afterEach, beforeAll, describe, expect, it } from "vitest";
import {
  applyFillKey,
  brokenIn,
  fillKey,
  hostSupports,
  numberFormat,
  overCap,
  pickScannableSheets,
  SELECTION_CELL_CAP,
  writeRuns,
} from "../src/excel/internal";
import { ANCHOR_PREFIX } from "../src/link/model";
import {
  currencyNumberFormat,
  getActiveSettings,
  setActiveSettings,
} from "../src/settings";

describe("overCap", () => {
  it("is false at zero and at the cap itself", () => {
    expect(overCap(0)).toBe(false);
    expect(overCap(SELECTION_CELL_CAP)).toBe(false);
  });

  it("is true one cell past the cap", () => {
    expect(overCap(SELECTION_CELL_CAP + 1)).toBe(true);
  });

  it("treats a negative count as over the cap - Excel's own answer past 2^31-1", () => {
    expect(overCap(-1)).toBe(true);
  });

  // The multi-area sum is proven where it happens: test/multi-area.integration.test.ts
  // (3 000 + 3 000 cells refused); overCap itself takes one scalar.
});

function fakeSheet(name: string): Excel.Worksheet {
  return { name } as unknown as Excel.Worksheet;
}

// pickScannableSheets trusts its ranges to arrive already loaded (isNullObject
// and cellCount synced by the caller), so a plain object of just those two
// fields is the whole fake a range needs to be here.
function fakeRange(cellCount: number, isNullObject = false): Excel.Range {
  return { isNullObject, cellCount } as unknown as Excel.Range;
}

describe("pickScannableSheets", () => {
  it("scans a normal sheet and a hidden one alike, skips one over the per-sheet cap, drops a negative count, and leaves an empty sheet off both lists", () => {
    const perSheetCap = 1_000;
    const sheets = [
      fakeSheet("Model"),
      // Hidden on purpose: this helper carries no visibility rule of its own
      // - that choice belongs to model-check.ts, which reads a hidden sheet
      // but never activates it. Here it scans like any other.
      fakeSheet("Assumptions (hidden)"),
      fakeSheet("Huge"),
      fakeSheet("Overflowed"),
      fakeSheet("Empty"),
    ];
    const ranges = [
      fakeRange(500),
      fakeRange(200),
      fakeRange(perSheetCap + 1),
      fakeRange(-1), // Excel's own answer for a count past 2^31-1
      fakeRange(0, true), // getUsedRangeOrNullObject found nothing at all
    ];

    const { scanned, skippedSheets } = pickScannableSheets(
      sheets,
      ranges,
      perSheetCap,
    );

    expect(scanned.map((s) => s.name)).toEqual([
      "Model",
      "Assumptions (hidden)",
    ]);
    expect(skippedSheets).toEqual(["Huge", "Overflowed"]);
  });

  it("skips only the sheet whose own addition would cross the workbook-wide scan cap", () => {
    const sheets = [fakeSheet("A"), fakeSheet("B"), fakeSheet("C")];
    const perSheetCap = 1_000; // generous: nothing here is skipped for its own size
    const totalCap = 150;
    const ranges = [fakeRange(100), fakeRange(60), fakeRange(10)];

    const { scanned, skippedSheets } = pickScannableSheets(
      sheets,
      ranges,
      perSheetCap,
      totalCap,
    );

    // A (100) fits under 150. B (60) would bring the running total to 160,
    // over the cap, so it is skipped WITHOUT joining the total. C (10) is
    // then asked against the total A alone left behind (100 + 10 = 110),
    // which still fits.
    expect(scanned.map((s) => s.name)).toEqual(["A", "C"]);
    expect(skippedSheets).toEqual(["B"]);
  });
});

describe("numberFormat", () => {
  it("returns the static whole, decimal and percent formats untouched by settings", () => {
    expect(numberFormat("whole")).toBe("#,##0;[Red](#,##0);-");
    expect(numberFormat("decimal")).toBe("#,##0.0;[Red](#,##0.0);-");
    expect(numberFormat("percent")).toBe("0.0%;[Red](0.0%);-");
  });

  it("builds the currency format from the active settings' own symbol and language", () => {
    const before = getActiveSettings();
    setActiveSettings({ ...before, currency: "$", language: "en" });
    try {
      expect(numberFormat("currency")).toBe(currencyNumberFormat("$", "en"));
    } finally {
      setActiveSettings(before);
    }
  });
});

describe("hostSupports", () => {
  afterEach(() => {
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("answers true before Office.context exists - the pane's state before the host connects", () => {
    (globalThis as { Office?: object }).Office = {};
    expect(hostSupports("1.9")).toBe(true);
  });

  it("defers to Office.context.requirements once the host has connected", () => {
    (
      globalThis as {
        Office?: {
          context: {
            requirements: {
              isSetSupported: (set: string, version: string) => boolean;
            };
          };
        };
      }
    ).Office = {
      context: {
        requirements: { isSetSupported: (_set, version) => version === "1.9" },
      },
    };
    expect(hostSupports("1.9")).toBe(true);
    expect(hostSupports("1.99")).toBe(false);
  });
});

// fillKey reads Excel.FillPattern.none as a plain runtime value (the real
// object lives in Office.js, which no unit test loads); this is the minimal
// stand-in, scoped to this describe block only. Vitest isolates each test
// file's globals, so it cannot leak into another file.
describe("fillKey and applyFillKey", () => {
  beforeAll(() => {
    (
      globalThis as { Excel?: { FillPattern: { none: string; solid: string } } }
    ).Excel = {
      FillPattern: { none: "None", solid: "Solid" },
    };
  });
  afterAll(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("reads no fill as the sentinel key, and applying it clears the fill", () => {
    const fill = { pattern: "None" } as unknown as Excel.CellPropertiesFill;
    expect(fillKey(fill)).toBe("none");

    let cleared = false;
    const block = {
      format: {
        fill: {
          clear: () => {
            cleared = true;
          },
        },
      },
    } as unknown as Excel.Range;
    applyFillKey(block, "none");
    expect(cleared).toBe(true);
  });

  it("round-trips a solid fill's pattern, colour and pattern colour through the key", () => {
    const fill = {
      pattern: "Solid",
      color: "#FF0000",
      patternColor: "#00FF00",
    } as unknown as Excel.CellPropertiesFill;
    const key = fillKey(fill);
    expect(key).toBe("Solid|#FF0000|#00FF00");

    const written: { color?: string; pattern?: string; patternColor?: string } =
      {};
    const block = {
      format: {
        fill: {
          clear: () => undefined,
          set color(value: string) {
            written.color = value;
          },
          set pattern(value: string) {
            written.pattern = value;
          },
          set patternColor(value: string) {
            written.patternColor = value;
          },
        },
      },
    } as unknown as Excel.Range;
    applyFillKey(block, key);
    expect(written).toEqual({
      color: "#FF0000",
      pattern: "Solid",
      patternColor: "#00FF00",
    });
  });

  it("falls back to white for a missing colour or pattern colour", () => {
    // No "|" at all: split() gives only one element, so the second and third
    // (colour, pattern colour) are genuinely undefined, not empty strings.
    const key = "Solid";
    const written: { color?: string; patternColor?: string } = {};
    const block = {
      format: {
        fill: {
          clear: () => undefined,
          set color(value: string) {
            written.color = value;
          },
          set pattern(_value: string) {
            /* not asserted here */
          },
          set patternColor(value: string) {
            written.patternColor = value;
          },
        },
      },
    } as unknown as Excel.Range;
    applyFillKey(block, key);
    expect(written).toEqual({ color: "#FFFFFF", patternColor: "#FFFFFF" });
  });
});

describe("brokenIn", () => {
  function names(
    items: { name: string; formula: unknown }[],
  ): Excel.NamedItemCollection {
    return { items } as unknown as Excel.NamedItemCollection;
  }

  it("flags a name whose formula collapsed to #REF! after the range it named was deleted", () => {
    const collection = names([
      { name: "TaxRate", formula: "=#REF!" },
      { name: "Revenue", formula: "=Model!$B$2" },
    ]);
    expect(brokenIn(collection)).toEqual(["TaxRate"]);
  });

  it("never flags a link anchor, even one left pointing at #REF!", () => {
    const collection = names([
      { name: `${ANCHOR_PREFIX}abcd1234`, formula: "=#REF!" },
    ]);
    expect(brokenIn(collection)).toEqual([]);
  });

  it("treats a non-string formula (a simple named constant) as never broken", () => {
    const collection = names([{ name: "Pi", formula: 3.14159 }]);
    expect(brokenIn(collection)).toEqual([]);
  });
});

describe("writeRuns", () => {
  interface Block {
    row: number;
    column: number;
    width: number;
  }

  // A minimal stand-in for the one call chain writeRuns makes on a range:
  // getCell(row, col).getResizedRange(0, width - 1). No load, no sync - the
  // grouping logic is the only thing under test.
  function fakeRange(): Excel.Range {
    return {
      getCell(row: number, column: number) {
        return {
          getResizedRange: (
            _deltaRows: number,
            deltaColumns: number,
          ): Block => ({
            row,
            column,
            width: deltaColumns + 1,
          }),
        };
      },
    } as unknown as Excel.Range;
  }

  it("writes one run per contiguous stretch of the same key, skipping null cells", () => {
    const written: { key: string; block: Block }[] = [];
    writeRuns(fakeRange(), [["a", "a", null, "b", "b", "b"]], (block, key) => {
      written.push({ key, block: block as unknown as Block });
    });
    expect(written).toEqual([
      { key: "a", block: { row: 0, column: 0, width: 2 } },
      { key: "b", block: { row: 0, column: 3, width: 3 } },
    ]);
  });

  it("keeps rows independent: a run never crosses from one row into the next", () => {
    const written: Block[] = [];
    writeRuns(
      fakeRange(),
      [
        ["x", "x"],
        ["x", "x"],
      ],
      (block) => {
        written.push(block as unknown as Block);
      },
    );
    expect(written).toEqual([
      { row: 0, column: 0, width: 2 },
      { row: 1, column: 0, width: 2 },
    ]);
  });
});
