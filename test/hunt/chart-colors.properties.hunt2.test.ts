// Pass-2 property: src/chart-colors.ts's brand rules. seriesPalette always
// hands back six usable colours; waterfallColors follows the
// totals/rises/falls rule for any sign pattern, zero included.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  pieColors,
  seriesPalette,
  waterfallColors,
  type WaterfallBrand,
} from "../../src/chart-colors";

const SEED = 20260927;
const RUNS = 3000;
const HEX = /^#[0-9A-F]{6}$/i;

const hexArb = fc
  .array(fc.integer({ min: 0, max: 255 }), { minLength: 3, maxLength: 3 })
  .map(
    ([r, g, b]) =>
      `#${[r, g, b].map((c) => c!.toString(16).padStart(2, "0")).join("")}`,
  );

const brandArb = fc.record({ primary: hexArb, accent: hexArb });
const waterfallBrandArb: fc.Arbitrary<WaterfallBrand> = fc.record({
  primary: hexArb,
  accent: hexArb,
  external: hexArb,
});

describe("seriesPalette: six usable colours for any brand", () => {
  it("always returns exactly six well-formed hex colours", () => {
    fc.assert(
      fc.property(brandArb, (brand) => {
        const palette = seriesPalette(brand);
        expect(palette).toHaveLength(6);
        for (const color of palette) expect(color).toMatch(HEX);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("gives every series a defined colour when cycled by index, any series count", () => {
    fc.assert(
      fc.property(
        brandArb,
        fc.integer({ min: 1, max: 40 }),
        (brand, seriesCount) => {
          const palette = seriesPalette(brand);
          for (let index = 0; index < seriesCount; index += 1) {
            const color = palette[index % palette.length];
            expect(color).toMatch(HEX);
          }
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("pieColors: one palette colour per slice, cycled past the sixth", () => {
  it("returns exactly `count` colours, each one straight from the palette", () => {
    fc.assert(
      fc.property(brandArb, fc.integer({ min: 0, max: 40 }), (brand, count) => {
        const palette = seriesPalette(brand);
        const colors = pieColors(count, brand);
        expect(colors).toHaveLength(count);
        colors.forEach((color, index) => {
          expect(color).toBe(palette[index % palette.length]);
        });
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("waterfallColors: totals primary, a fall external, a rise accent", () => {
  const amountArb = fc
    .tuple(
      fc.boolean(),
      fc.double({ min: 0, max: 10, noNaN: true }),
      fc.integer({ min: -6, max: 9 }),
    )
    .map(([negative, mantissa, exponent]) => {
      const value = mantissa * 10 ** exponent;
      return negative ? -value : value;
    });

  it("colours every point for any sign pattern, opening/closing always primary", () => {
    fc.assert(
      fc.property(
        fc.array(amountArb, { minLength: 3, maxLength: 30 }),
        waterfallBrandArb,
        (values, brand) => {
          const colors = waterfallColors(values, brand);
          expect(colors).toHaveLength(values.length);
          expect(colors[0]).toBe(brand.primary);
          expect(colors[colors.length - 1]).toBe(brand.primary);
          for (let index = 1; index < values.length - 1; index += 1) {
            const expected = values[index]! < 0 ? brand.external : brand.accent;
            expect(colors[index]).toBe(expected);
          }
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("colours two points or fewer entirely primary, opening and closing with nothing between", () => {
    fc.assert(
      fc.property(
        fc.array(amountArb, { minLength: 0, maxLength: 2 }),
        waterfallBrandArb,
        (values, brand) => {
          const colors = waterfallColors(values, brand);
          expect(colors).toEqual(values.map(() => brand.primary));
        },
      ),
      { seed: SEED, numRuns: 500 },
    );
  });

  it("treats an exact zero delta as a rise (accent), never as a fall", () => {
    fc.assert(
      fc.property(
        fc.array(amountArb, { minLength: 1, maxLength: 15 }),
        fc.array(amountArb, { minLength: 1, maxLength: 15 }),
        waterfallBrandArb,
        (before, after, brand) => {
          const values = [...before, 0, ...after];
          const index = before.length;
          const colors = waterfallColors(values, brand);
          expect(colors[index]).toBe(brand.accent);
        },
      ),
      { seed: SEED, numRuns: 500 },
    );
  });
});
