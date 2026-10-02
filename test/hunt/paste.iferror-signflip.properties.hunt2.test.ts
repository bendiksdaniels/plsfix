// Pass-2 property: toggleIfError and flipSign are each meant to undo their own
// wrap on a second press, and countIfErrorToggle is meant to agree with what
// toggleIfError actually did. Hammered with fast-check over a small formula
// grammar: numbers, refs, quoted/unquoted sheet names (commas, parens,
// apostrophes, unicode), table refs (brackets, ' escapes), string literals
// (commas, parens, doubled quotes), nested calls and operators - the exact
// syntax eachSyntaxChar (src/paste.ts) has to see through in either direction.
import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { isFormula } from "../../src/model";
import { countIfErrorToggle, flipSign, toggleIfError } from "../../src/paste";
import { sheetPrefix } from "../../src/formula-refs";

const SEED = 20260927;
const RUNS = 2000;

// Independent oracle: are this string's parens balanced, counting only the
// characters that are formula syntax outside a "..." string, a '...' sheet
// name (a doubled apostrophe re-opens it) and a [...] table body (where '
// escapes the next character)? Mirrors eachSyntaxChar's own lexical rules,
// which ARE the definition of "syntax" here, not a copy of its decision.
function hasBalancedSyntax(body: string): boolean {
  let depth = 0;
  let quote: string | null = null;
  let bracket = 0;
  for (let index = 0; index < body.length; index += 1) {
    const char = body[index];
    if (quote !== null) {
      if (char === quote) quote = null;
      continue;
    }
    if (bracket > 0) {
      if (char === "'") {
        index += 1;
      } else if (char === "[") {
        bracket += 1;
      } else if (char === "]") {
        bracket -= 1;
      }
      continue;
    }
    if (char === '"' || char === "'") quote = char;
    else if (char === "[") bracket += 1;
    else if (char === "(") depth += 1;
    else if (char === ")") {
      depth -= 1;
      if (depth < 0) return false;
    }
  }
  return depth === 0 && quote === null && bracket === 0;
}

// Sheet-qualified refs built through the real sheetPrefix, so every quoting
// rule (apostrophe doubling, commas, parens, unicode) is exactly what Excel
// would write - the same helper formula-refs.properties.hunt2 round-trips.
const trickyNames = [
  "Sales",
  "Sales, EU",
  "O'Brien",
  "Q1 (EU)",
  "O''Brien, Ltd",
  "P&L (2025)",
  "Pārskats",
  "销售数据",
];
const sheetRefAtom = fc
  .tuple(fc.constantFrom(...trickyNames), fc.constantFrom("A1", "$B$2", "C3"))
  .map(([name, cell]) => sheetPrefix(name) + cell);

const tableRefAtom = fc.constantFrom(
  "Table1[Col]",
  "Table1[[#This Row],[Price]]",
  "T[[#This Row],[a']b]]",
  "Table1[[#Headers],[A1]]",
  "T2[[#Data],[Q1, EU]]",
);

const stringLiteralAtom = fc.constantFrom(
  '"a,b"',
  '"x""y,z""w"',
  '"(paren)"',
  '"n/a"',
  '""',
  '"comma, and (parens)"',
);

const refAtom = fc.constantFrom("A1", "B2", "$C$3", "Z9", "AA10");
const numberAtom = fc.constantFrom("1", "0", "-3.5", "1000000", "0.001");

const atom = fc.oneof(
  refAtom,
  numberAtom,
  sheetRefAtom,
  tableRefAtom,
  stringLiteralAtom,
);

const OPS = ["+", "-", "*", "/", "&"];
// IFERROR is deliberately never a pickable function name here: toggleIfError
// treats ANY top-level IFERROR(...) call as already guarded regardless of who
// wrote it (pinned in paste.test.ts/paste.audit.test.ts), so a body that is
// one by construction is a different, already-covered scenario, not this
// round trip. flipSign's "-(" root shape gets the same treatment below.
const FUNCS = ["SUM", "VLOOKUP", "IF", "MAX", "MIN", "ROUND", "CONCATENATE"];

const { expr } = fc.letrec((tie) => ({
  expr: fc.oneof(
    { weight: 3, arbitrary: atom },
    {
      weight: 2,
      arbitrary: fc
        .tuple(tie("expr"), fc.constantFrom(...OPS), tie("expr"))
        .map(([a, op, b]) => `${a}${op}${b}`),
    },
    {
      weight: 1,
      arbitrary: fc
        .tuple(
          fc.constantFrom(...FUNCS),
          fc.array(tie("expr") as fc.Arbitrary<string>, {
            minLength: 1,
            maxLength: 3,
          }),
        )
        .map(([fn, args]) => `${fn}(${args.join(",")})`),
    },
    {
      weight: 1,
      arbitrary: (tie("expr") as fc.Arbitrary<string>).map((e) => `(${e})`),
    },
    {
      weight: 1,
      arbitrary: (tie("expr") as fc.Arbitrary<string>).map((e) => `-${e}`),
    },
  ),
}));

// The ROOT is never a bare unary-minus or paren-group and never a call to a
// name outside FUNCS, so the whole rendered string can never itself start
// with "IFERROR(" or with "-(" spanning to the very end - the two shapes
// toggleIfError/flipSign treat as "already wrapped, strip it" rather than
// "wrap it". first is always an atom or a named call; rest can be anything,
// including unary-minus or parens, just never as the very first character.
const rootExpr: fc.Arbitrary<string> = fc
  .tuple(
    fc.oneof(
      atom,
      fc
        .tuple(
          fc.constantFrom(...FUNCS),
          fc.array(expr, { minLength: 1, maxLength: 3 }),
        )
        .map(([fn, args]) => `${fn}(${args.join(",")})`),
    ),
    fc.array(fc.tuple(fc.constantFrom(...OPS), expr), {
      minLength: 0,
      maxLength: 2,
    }),
  )
  .map(([first, rest]) => first + rest.map(([op, e]) => op + e).join(""));

const fallbackArb = fc.constantFrom("0", '""', '"n/a"', "-1");

describe("toggleIfError twice returns the input", () => {
  it("wraps then unwraps back to the exact original body, for a fixed fallback", () => {
    fc.assert(
      fc.property(rootExpr, fallbackArb, (body, fallback) => {
        const cells = [[`=${body}`]];
        const once = toggleIfError(cells, fallback);
        const twice = toggleIfError(once, fallback);
        expect(twice).toEqual(cells);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("keeps every output's parens balanced outside strings, names and brackets", () => {
    fc.assert(
      fc.property(rootExpr, fallbackArb, (body, fallback) => {
        const once = toggleIfError([[`=${body}`]], fallback)[0]![0];
        expect(hasBalancedSyntax(String(once).slice(1))).toBe(true);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("countIfErrorToggle agrees with what toggleIfError did", () => {
  it("counts an add exactly where the output is the wrap of the input, a strip otherwise", () => {
    fc.assert(
      fc.property(
        fc.array(rootExpr, { minLength: 1, maxLength: 5 }),
        fc.array(fc.boolean(), { minLength: 1, maxLength: 5 }),
        fallbackArb,
        (bodies, guardedFlags, fallback) => {
          // Half the generated cells start already guarded (constructed
          // directly, not by a prior toggle - the same shape a modeller who
          // typed IFERROR by hand would leave), so this exercises both of
          // toggleIfError's own branches, not only the wrap direction.
          const row = bodies.map((body, index) =>
            guardedFlags[index % guardedFlags.length]
              ? `=IFERROR(${body},${fallback})`
              : `=${body}`,
          );
          const grid = [row];

          const toggled = toggleIfError(grid, fallback)[0]!;
          let expectedAdded = 0;
          let expectedStripped = 0;
          row.forEach((before, index) => {
            const after = toggled[index];
            if (!isFormula(before)) return;
            const wrapShape = `=IFERROR(${before.slice(1)},${fallback})`;
            if (after === wrapShape) expectedAdded += 1;
            else expectedStripped += 1;
          });

          expect(countIfErrorToggle(grid)).toEqual({
            added: expectedAdded,
            stripped: expectedStripped,
          });
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("flipSign twice returns the input, or an arithmetically equal formula", () => {
  it("wraps then unwraps back to the exact original body, for any root expression", () => {
    fc.assert(
      fc.property(rootExpr, (body) => {
        const cells = [[`=${body}`]];
        const twice = flipSign(flipSign(cells));
        expect(twice).toEqual(cells);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("is an exact involution on plain numbers, including both zeros", () => {
    // Negation flips the sign bit; flipping it twice always restores the
    // exact original bit pattern, -0 included - so the input comes back
    // byte for byte, never normalized to +0 along the way.
    fc.assert(
      fc.property(
        fc.oneof(fc.double({ noNaN: true }), fc.constant(0), fc.constant(-0)),
        (value) => {
          const twice = flipSign(flipSign([[value]]))[0]![0];
          expect(Object.is(twice, value)).toBe(true);
        },
      ),
      { seed: SEED, numRuns: 300 },
    );
  });

  // A formula the modeller wrote with its own redundant double negation
  // ("=-(-(A1))") is the one documented case flipSign does not restore byte
  // for byte on a second press: each press manages exactly one "-( ... )"
  // layer, so pressing twice sheds BOTH of the user's own layers instead of
  // restoring them - "=A1", not "=-(-(A1))". That is still arithmetically
  // the same sign (both read +A1), which is the guarantee this pins by name.
  it("sheds a user's own redundant double negation, but keeps the sign", () => {
    expect(flipSign([["=-(-(A1))"]])).toEqual([["=-(A1)"]]);
    expect(flipSign(flipSign([["=-(-(A1))"]]))).toEqual([["=A1"]]);
    // +A1 (double negative) in, +A1 (bare) out: the sign survived two presses.
  });

  it("keeps every output's parens balanced outside strings, names and brackets", () => {
    fc.assert(
      fc.property(rootExpr, (body) => {
        const once = flipSign([[`=${body}`]])[0]![0];
        expect(hasBalancedSyntax(String(once).slice(1))).toBe(true);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
