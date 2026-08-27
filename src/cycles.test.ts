import { describe, expect, it } from "vitest";
import {
  buildFillCycle,
  buildFontCycle,
  buildNumberCycles,
  buildRowStyleCycles,
  CLEAR_FILL,
  matchStyleIndex,
  nextInCycle,
  type StyleSpec,
} from "./cycles";
import {
  currencyNumberFormat,
  DEFAULT_SETTINGS,
  deriveTheme,
} from "./settings";

const theme = deriveTheme(DEFAULT_SETTINGS);

describe("number format cycles", () => {
  it("starts every family on the format the buttons already apply", () => {
    const cycles = buildNumberCycles(DEFAULT_SETTINGS);
    expect(cycles.general[0]).toBe("#,##0;[Red](#,##0);-");
    expect(cycles.general[1]).toBe("#,##0.0;[Red](#,##0.0);-");
    expect(cycles.percent[0]).toBe("0.0%;[Red](0.0%);-");
    expect(cycles.currency[0]).toBe(currencyNumberFormat(DEFAULT_SETTINGS.currency));
  });

  it("steps general through whole, one and two decimals", () => {
    expect(buildNumberCycles(DEFAULT_SETTINGS).general).toEqual([
      "#,##0;[Red](#,##0);-",
      "#,##0.0;[Red](#,##0.0);-",
      "#,##0.00;[Red](#,##0.00);-",
    ]);
  });

  it("follows the configured currency symbol", () => {
    const dollar = buildNumberCycles({ ...DEFAULT_SETTINGS, currency: "$" });
    expect(dollar.currency).toEqual([
      "$ #,##0;[Red]($ #,##0);-",
      "$ #,##0.0;[Red]($ #,##0.0);-",
      "$ #,##0,;[Red]($ #,##0,);-",
    ]);

    const bare = buildNumberCycles({ ...DEFAULT_SETTINGS, currency: "" });
    expect(bare.currency[0]).toBe("#,##0;[Red](#,##0);-");
    expect(bare.currency[2]).toBe("#,##0,;[Red](#,##0,);-");
  });

  it("cycles percent, multiple and date families", () => {
    const cycles = buildNumberCycles(DEFAULT_SETTINGS);
    expect(cycles.percent).toEqual([
      "0.0%;[Red](0.0%);-",
      "0%;[Red](0%);-",
      "0.00%;[Red](0.00%);-",
    ]);
    expect(cycles.multiple).toEqual([
      '0.0"x";[Red](0.0"x");-',
      '0.00"x";[Red](0.00"x");-',
    ]);
    expect(cycles.date).toEqual(["dd.mm.yyyy", "mmm-yy", "yyyy"]);
  });
});

describe("cycle stepping", () => {
  const cycle = ["a", "b", "c"];

  it("steps to the next entry and wraps at the end", () => {
    expect(nextInCycle("a", cycle)).toBe("b");
    expect(nextInCycle("b", cycle)).toBe("c");
    expect(nextInCycle("c", cycle)).toBe("a");
  });

  it("starts the cycle when the current format is not one of ours", () => {
    expect(nextInCycle("General", cycle)).toBe("a");
    expect(nextInCycle("", cycle)).toBe("a");
  });

  it("leaves the cell alone when the cycle is empty", () => {
    expect(nextInCycle("General", [])).toBe("General");
  });
});

describe("row style cycles", () => {
  const cycles = buildRowStyleCycles(DEFAULT_SETTINGS);

  it("starts the title cycle on the solid brand fill", () => {
    expect(cycles.title[0]).toEqual({
      fill: theme.titleFill,
      fontColor: theme.titleText,
      bold: true,
      topBorder: null,
      bottomBorder: null,
    });
  });

  it("starts the result cycle on the shipped result preset", () => {
    expect(cycles.result[0]).toEqual({
      fill: theme.resultFill,
      fontColor: theme.formulaFont,
      bold: true,
      topBorder: { style: "double", color: theme.resultBorder },
      bottomBorder: null,
    });
  });

  it("starts the item cycle on a plain unfilled row", () => {
    expect(cycles.item[0]).toEqual({
      fill: CLEAR_FILL,
      fontColor: theme.formulaFont,
      bold: false,
      topBorder: null,
      bottomBorder: null,
    });
  });

  it("keeps every cycle between two and three variants", () => {
    for (const variants of Object.values(cycles)) {
      expect(variants.length).toBeGreaterThanOrEqual(2);
      expect(variants.length).toBeLessThanOrEqual(3);
    }
  });

  it("derives every variant color from the palette", () => {
    const palette = new Set<string>([
      CLEAR_FILL,
      ...Object.values(theme),
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.accent,
    ]);

    for (const variants of Object.values(cycles)) {
      for (const variant of variants) {
        if (variant.fill) expect(palette.has(variant.fill)).toBe(true);
        if (variant.fontColor) expect(palette.has(variant.fontColor)).toBe(true);
        if (variant.topBorder) expect(palette.has(variant.topBorder.color)).toBe(true);
        if (variant.bottomBorder) expect(palette.has(variant.bottomBorder.color)).toBe(true);
      }
    }
  });
});

describe("style matching", () => {
  const variants: StyleSpec[] = [
    { fill: "#282623", fontColor: "#FFFFFF", bold: true },
    { fill: CLEAR_FILL, bold: false },
    { bold: true },
  ];

  it("matches on the properties a variant defines", () => {
    expect(
      matchStyleIndex({ fill: "#282623", fontColor: "#FFFFFF", bold: true }, variants),
    ).toBe(0);
    expect(
      matchStyleIndex({ fill: null, fontColor: "#1F1D1B", bold: false }, variants),
    ).toBe(1);
    expect(
      matchStyleIndex({ fill: "#B27E54", fontColor: "#1F1D1B", bold: true }, variants),
    ).toBe(2);
  });

  it("ignores font color when the variant does not set one", () => {
    expect(
      matchStyleIndex({ fill: null, fontColor: "#B27E54", bold: false }, variants),
    ).toBe(1);
  });

  it("returns -1 when nothing matches", () => {
    expect(
      matchStyleIndex({ fill: "#B27E54", fontColor: "#1F1D1B", bold: false }, [
        variants[0] as StyleSpec,
      ]),
    ).toBe(-1);
  });
});

describe("fill and font cycles", () => {
  it("cycles brand fills and ends on a cleared cell", () => {
    expect(buildFillCycle(DEFAULT_SETTINGS)).toEqual([
      theme.headerFill,
      theme.resultFill,
      DEFAULT_SETTINGS.accent,
      DEFAULT_SETTINGS.primary,
      CLEAR_FILL,
    ]);
  });

  it("cycles the model font colors", () => {
    expect(buildFontCycle(DEFAULT_SETTINGS)).toEqual([
      theme.formulaFont,
      theme.inputFont,
      theme.linkFont,
      DEFAULT_SETTINGS.accent,
      DEFAULT_SETTINGS.primary,
    ]);
  });
});
