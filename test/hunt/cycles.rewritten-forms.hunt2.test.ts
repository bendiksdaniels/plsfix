// Property: every format cycle in src/cycles.ts steps to the true next state
// whatever rewritten form Excel hands back for the current one: locale-tagged
// currency, a bare locale tag, "Accountant" underlines, "General" in any case,
// border readouts in any case.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  ALIGN_CYCLE,
  buildBorderCycle,
  buildNumberCycles,
  buildSizeCycles,
  canonicalNumberFormat,
  matchBorderIndex,
  nextAlignment,
  nextInCycle,
  nextSize,
  nextUnderline,
  UNDERLINE_CYCLE,
  type BorderCycleState,
  type BorderReadouts,
  type NumberCycleFamily,
} from "../../src/cycles";
import { DEFAULT_SETTINGS } from "../../src/settings";

const SEED = 20260927;
const RUNS = 2000;

// ---------------------------------------------------------------------------
// Number format families: a currency symbol wrapped the way Excel tags it on
// read-back, or a bare locale tag over a format with no symbol at all.
// ---------------------------------------------------------------------------

const NUMBER_CYCLES = buildNumberCycles(DEFAULT_SETTINGS);
const FAMILIES = Object.keys(NUMBER_CYCLES) as NumberCycleFamily[];
const LOCALE_TAGS = ["409", "422", "x-euro2", "07", ""];

type NumberRewrite = "identity" | "bareTag" | "wrapSymbol";

function rewriteNumberFormat(
  canonical: string,
  style: NumberRewrite,
  tag: string,
  symbol: string,
): string {
  if (style === "identity") return canonical;
  if (style === "bareTag") return `[$-${tag}]${canonical}`;
  if (!symbol || !canonical.includes(symbol)) return canonical;
  return canonical.replace(symbol, `[$${symbol}-${tag}]`);
}

const familyArb = fc.constantFrom(...FAMILIES);
const rewriteStyleArb = fc.constantFrom<NumberRewrite>(
  "identity",
  "bareTag",
  "wrapSymbol",
);
const tagArb = fc.constantFrom(...LOCALE_TAGS);

describe("number format cycles step past any rewritten form of the current state", () => {
  it("lands on the true next entry for every family, index and rewrite", () => {
    fc.assert(
      fc.property(
        familyArb,
        rewriteStyleArb,
        tagArb,
        fc.nat({ max: 2 }),
        (family, style, tag, offset) => {
          const cycle = NUMBER_CYCLES[family];
          const index = offset % cycle.length;
          const current = cycle[index]!;
          const rewritten = rewriteNumberFormat(
            current,
            style,
            tag,
            DEFAULT_SETTINGS.currency,
          );
          // The rewrite must still canonicalize back to the same state, or
          // the property would be asking the wrong question.
          expect(canonicalNumberFormat(rewritten)).toBe(
            canonicalNumberFormat(current),
          );
          const expectedNext = cycle[(index + 1) % cycle.length];
          expect(nextInCycle(rewritten, cycle)).toBe(expectedNext);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

// ---------------------------------------------------------------------------
// Alignment and underline: Excel's own case, plus "SingleAccountant" /
// "DoubleAccountant" for the accounting underline rungs.
// ---------------------------------------------------------------------------

function rewriteCase(
  value: string,
  style: "lower" | "upper" | "as-is",
): string {
  if (style === "lower") return value.toLowerCase();
  if (style === "upper") return value.toUpperCase();
  return value;
}

const caseStyleArb = fc.constantFrom<"lower" | "upper" | "as-is">(
  "lower",
  "upper",
  "as-is",
);

describe("nextAlignment steps past any case of the rung it is already on", () => {
  it("includes General, whatever case it is spelled in", () => {
    fc.assert(
      fc.property(
        fc.nat({ max: ALIGN_CYCLE.length - 1 }),
        caseStyleArb,
        (index, style) => {
          const current = rewriteCase(ALIGN_CYCLE[index]!, style);
          const expected = ALIGN_CYCLE[(index + 1) % ALIGN_CYCLE.length];
          expect(nextAlignment(current)).toBe(expected);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("nextUnderline steps past any case, Accountant suffix included", () => {
  it("treats SingleAccountant/DoubleAccountant as the plain rung", () => {
    fc.assert(
      fc.property(
        fc.nat({ max: UNDERLINE_CYCLE.length - 1 }),
        caseStyleArb,
        fc.boolean(),
        (index, style, accountant) => {
          const rung = UNDERLINE_CYCLE[index]!;
          // Excel only ever answers "Accountant" behind Single or Double,
          // never behind None: a book-keeping underline is always one or
          // the other.
          const suffixed =
            accountant && rung !== "None" ? `${rung}Accountant` : rung;
          const current = rewriteCase(suffixed, style);
          const expected =
            UNDERLINE_CYCLE[(index + 1) % UNDERLINE_CYCLE.length];
          expect(nextUnderline(current)).toBe(expected);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

// ---------------------------------------------------------------------------
// Border cycle: a readout in any case for style, weight and the colour's hex
// digits must still match the state that drew it, so the cycle can step past
// it - matchBorderIndex must find the SAME index whatever case the host used.
// ---------------------------------------------------------------------------

const BORDER_STATES: BorderCycleState[] = buildBorderCycle(DEFAULT_SETTINGS);

function rewriteHexCase(
  hex: string,
  style: "lower" | "upper" | "as-is",
): string {
  return rewriteCase(hex, style);
}

// The readout matchBorderIndex would see for a state fully drawn as itself,
// then rewritten in the given case for style, weight and colour.
function readoutFor(
  state: BorderCycleState,
  style: "lower" | "upper" | "as-is",
): BorderReadouts {
  const readouts: BorderReadouts = {};
  for (const edge of [
    "top",
    "bottom",
    "left",
    "right",
    "insideHorizontal",
    "insideVertical",
  ] as const) {
    const line = state.find((entry) => entry.edge === edge);
    readouts[edge] = line
      ? {
          style: rewriteCase(
            line.style === "continuous" ? "Continuous" : "Double",
            style,
          ),
          weight: rewriteCase(
            line.weight === "thin" ? "Thin" : "Medium",
            style,
          ),
          color: rewriteHexCase(line.color, style),
        }
      : { style: rewriteCase("None", style), weight: "Thin", color: "#000000" };
  }
  return readouts;
}

describe("matchBorderIndex finds the same state whatever case the host answers in", () => {
  it("over every border cycle state and every edge's case", () => {
    fc.assert(
      fc.property(
        fc.nat({ max: BORDER_STATES.length - 1 }),
        caseStyleArb,
        (index, style) => {
          const readout = readoutFor(BORDER_STATES[index]!, style);
          expect(matchBorderIndex(readout, BORDER_STATES)).toBe(index);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

// ---------------------------------------------------------------------------
// Size cycles: Excel snaps a written height or width to the nearest screen
// pixel, so a read-back within nextSize's own half-point tolerance must still
// count as the rung that was written.
// ---------------------------------------------------------------------------

const SIZE_CYCLES = buildSizeCycles();

describe("nextSize steps past a pixel-snapped read-back of the current rung", () => {
  it("for every rung of both size cycles, within the documented tolerance", () => {
    fc.assert(
      fc.property(
        fc.constantFrom<"rowHeight" | "columnWidth">(
          "rowHeight",
          "columnWidth",
        ),
        fc.nat({ max: 4 }),
        fc.double({ min: -0.5, max: 0.5, noNaN: true }),
        (which, offset, drift) => {
          const cycle = SIZE_CYCLES[which];
          const index = offset % cycle.length;
          const current = cycle[index]! + drift;
          const expected = cycle[(index + 1) % cycle.length];
          expect(nextSize(current, cycle)).toBe(expected);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
