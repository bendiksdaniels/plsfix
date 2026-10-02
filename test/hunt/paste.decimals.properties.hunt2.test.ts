// Pass-2 property: stepDecimals up then down is the identity, and
// formatDecimals tracks it exactly, for generated number formats - sections
// separated by ";", quoted literals, backslash escapes, percent, currency
// tags like [$EUR-x-euro2] and a bare locale tag [$-409].
import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { formatDecimals, stepDecimals } from "../../src/paste";

const SEED = 20260927;
const RUNS = 3000;

const literalQuoted = fc.constantFrom('"kr"', '"x"', '" ; "', '"a,b"', '""');
const bracketTag = fc.constantFrom(
  "[$€-x-euro2]",
  "[$-409]",
  "[Red]",
  "[Blue]",
  "[$USD]",
);
const escapedChar = fc.constantFrom("\\ ", "\\%", "\\,", "\\.", "\\-");
const prefixPiece = fc.oneof(
  literalQuoted,
  bracketTag,
  escapedChar,
  fc.constant("€ "),
  fc.constant(""),
);

const digitChar = fc.constantFrom("0", "#", "?");
function digitsRun(min: number, max: number): fc.Arbitrary<string> {
  return fc
    .array(digitChar, { minLength: min, maxLength: max })
    .map((chars) => chars.join(""));
}

// intPart never empty: at least one placeholder, so numericBody always finds
// something to work with in the "with digits" branch. A dot, once present,
// always carries at least one placeholder after it - a bare trailing dot
// with none ("0.") is excluded here on purpose; see the dedicated test below
// for why that one specific shape cannot round-trip.
const numericCore = fc
  .tuple(
    digitsRun(1, 3),
    fc.oneof(
      fc.constant(""),
      digitsRun(1, 3).map((decPart) => `.${decPart}`),
    ),
    fc.integer({ min: 0, max: 2 }),
  )
  .map(([intPart, dot, commas]) => intPart + dot + ",".repeat(commas));

// A real percent (scales the printed value by 100) or an escaped, literal
// percent sign (does not) - both attached directly after the digits, the
// exact shape of the pinned "0.0\%" regression.
const percentSuffix = fc.constantFrom("", "%", "\\%");

const sectionWithDigits = fc
  .tuple(prefixPiece, numericCore, percentSuffix, prefixPiece)
  .map(([pre, core, pct, post]) => pre + core + pct + post);

const sectionNoDigits = fc.constantFrom("General", "@", "dd.mm.yyyy", "");

const section = fc.oneof(
  { weight: 4, arbitrary: sectionWithDigits },
  { weight: 1, arbitrary: sectionNoDigits },
);

const formatArb = fc
  .array(section, { minLength: 1, maxLength: 4 })
  .map((sections) => sections.join(";"));

describe("stepDecimals up then down is the identity", () => {
  it("returns the exact original format for every generated section family", () => {
    fc.assert(
      fc.property(formatArb, (format) => {
        expect(stepDecimals(stepDecimals(format, 1), -1)).toBe(format);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("never throws, whatever section shape it is handed", () => {
    fc.assert(
      fc.property(
        formatArb,
        fc.constantFrom(1, -1) as fc.Arbitrary<1 | -1>,
        (format, delta) => {
          expect(() => stepDecimals(format, delta)).not.toThrow();
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("formatDecimals tracks stepDecimals exactly", () => {
  it("adds one once stepped up, or stays null when there was nothing to step", () => {
    fc.assert(
      fc.property(formatArb, (format) => {
        const before = formatDecimals(format);
        const after = formatDecimals(stepDecimals(format, 1));
        if (before === null) {
          expect(after).toBeNull();
        } else {
          expect(after).toBe(before + 1);
        }
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("agrees with a fresh stepDecimals(format, -1) the same way, one below", () => {
    fc.assert(
      fc.property(formatArb, (format) => {
        const before = formatDecimals(format);
        const after = formatDecimals(stepDecimals(format, -1));
        if (before === null) {
          expect(after).toBeNull();
        } else {
          // stepSection's own floor: a format with no decimal point yet does
          // not go negative on a -1 step (src/paste.ts stepSection, the
          // "if (delta === -1) return section;" branch when dot < 0) - so
          // stepping down lands one below, or stays put with nothing to
          // remove. It never overshoots either way.
          expect([before, before - 1]).toContain(after);
        }
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

// Found by the round-trip property before the grammar above excluded it:
// "0." (a literal trailing decimal point with NO placeholder after it) is a
// legitimate, if unusual, number format - but stepDecimals(+1) on it inserts
// a placeholder right after the dot exactly the same way it would for "0"
// (no dot at all): both become "0.0". Once that happens the two starting
// shapes are indistinguishable, so stepDecimals(-1) cannot tell them apart
// either and always drops back to the dot-free "0" - matching Excel's own
// Decrease Decimal, which also drops the point once no placeholder is left.
// There is no smallest fix here: preserving "0."'s bare dot through the
// round trip would require "0" (the overwhelmingly more common shape) to
// stop round-tripping instead, since increase-decimal cannot leave a mark
// distinguishing which of the two it started from. Pinned as documented,
// not fixed.
describe("a bare trailing dot with no placeholder does not survive up-then-down", () => {
  it("converges with the dot-free shape once stepped up, and stays converged back down", () => {
    expect(stepDecimals("0.", 1)).toBe("0.0");
    expect(stepDecimals("0", 1)).toBe("0.0");
    expect(stepDecimals("0.0", -1)).toBe("0");
    expect(stepDecimals(stepDecimals("0.", 1), -1)).toBe("0");
  });
});
