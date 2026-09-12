import { describe, expect, it } from "vitest";
import {
  ALIGN_CYCLE,
  type BorderCycleState,
  type BorderEdgeName,
  type BorderReadouts,
  buildBorderCycle,
  buildFillCycle,
  buildFontCycle,
  buildNumberCycles,
  buildRowStyleCycles,
  buildSizeCycles,
  canonicalAlignment,
  canonicalNumberFormat,
  canonicalUnderline,
  CLEAR_FILL,
  INDENT_CYCLE,
  matchBorderIndex,
  matchStyleIndex,
  nextAlignment,
  nextIndent,
  nextInCycle,
  nextSize,
  nextUnderline,
  type StyleSpec,
  UNDERLINE_CYCLE,
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
    expect(cycles.currency[0]).toBe(
      currencyNumberFormat(
        DEFAULT_SETTINGS.currency,
        DEFAULT_SETTINGS.language,
      ),
    );
  });

  it("steps general through whole, one and two decimals", () => {
    expect(buildNumberCycles(DEFAULT_SETTINGS).general).toEqual([
      "#,##0;[Red](#,##0);-",
      "#,##0.0;[Red](#,##0.0);-",
      "#,##0.00;[Red](#,##0.00);-",
    ]);
  });

  it("puts the symbol after the amount in Latvian and Russian", () => {
    const lv = buildNumberCycles({ ...DEFAULT_SETTINGS, language: "lv" });
    expect(lv.currency).toEqual([
      "#,##0 €;[Red](#,##0 €);-",
      "#,##0.0 €;[Red](#,##0.0 €);-",
      "#,##0, €;[Red](#,##0, €);-",
    ]);
    expect(
      buildNumberCycles({ ...DEFAULT_SETTINGS, language: "ru" }).currency,
    ).toEqual(lv.currency);
  });

  it("follows the configured currency symbol", () => {
    const dollar = buildNumberCycles({
      ...DEFAULT_SETTINGS,
      language: "en",
      currency: "$",
    });
    expect(dollar.currency).toEqual([
      "$ #,##0;[Red]($ #,##0);-",
      "$ #,##0.0;[Red]($ #,##0.0);-",
      "$ #,##0,;[Red]($ #,##0,);-",
    ]);

    const bare = buildNumberCycles({
      ...DEFAULT_SETTINGS,
      language: "en",
      currency: "",
    });
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

  it("strips locale-tagged currency codes when canonicalizing", () => {
    expect(
      canonicalNumberFormat("[$€-x-euro2] #,##0;[Red]([$€-x-euro2] #,##0);-"),
    ).toBe("€ #,##0;[Red](€ #,##0);-");
    expect(canonicalNumberFormat("[$$-409] #,##0.0")).toBe("$ #,##0.0");
    expect(canonicalNumberFormat("#,##0;[Red](#,##0);-")).toBe(
      "#,##0;[Red](#,##0);-",
    );
  });

  it("keeps cycling when Excel rewrote the applied currency format", () => {
    const currency = buildNumberCycles(DEFAULT_SETTINGS).currency;
    const rewritten = "#,##0 [$€-x-euro2];[Red](#,##0 [$€-x-euro2]);-";
    expect(nextInCycle(rewritten, currency)).toBe(currency[1]);
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
        if (variant.fontColor)
          expect(palette.has(variant.fontColor)).toBe(true);
        if (variant.topBorder)
          expect(palette.has(variant.topBorder.color)).toBe(true);
        if (variant.bottomBorder)
          expect(palette.has(variant.bottomBorder.color)).toBe(true);
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
      matchStyleIndex(
        { fill: "#282623", fontColor: "#FFFFFF", bold: true },
        variants,
      ),
    ).toBe(0);
    expect(
      matchStyleIndex(
        { fill: null, fontColor: "#1F1D1B", bold: false },
        variants,
      ),
    ).toBe(1);
    expect(
      matchStyleIndex(
        { fill: "#B27E54", fontColor: "#1F1D1B", bold: true },
        variants,
      ),
    ).toBe(2);
  });

  it("ignores font color when the variant does not set one", () => {
    expect(
      matchStyleIndex(
        { fill: null, fontColor: "#B27E54", bold: false },
        variants,
      ),
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

describe("border cycle", () => {
  const states = buildBorderCycle(DEFAULT_SETTINGS);
  const line = DEFAULT_SETTINGS.primary;
  const OUTER: BorderEdgeName[] = ["top", "bottom", "left", "right"];
  const ALL: BorderEdgeName[] = [
    ...OUTER,
    "insideHorizontal",
    "insideVertical",
  ];

  // What the host hands back after a state is written: Excel's own spelling,
  // and "None" for every edge the state leaves alone.
  function readback(
    state: BorderCycleState,
    edges: BorderEdgeName[] = ALL,
  ): BorderReadouts {
    const out: BorderReadouts = {};
    for (const edge of edges) {
      const drawn = state.find((entry) => entry.edge === edge);
      out[edge] = drawn
        ? {
            style: drawn.style === "double" ? "Double" : "Continuous",
            weight: drawn.weight === "medium" ? "Medium" : "Thin",
            color: drawn.color,
          }
        : { style: "None", weight: "Thin", color: "#000000" };
    }
    return out;
  }

  it("runs from nothing to a full grid in six looks", () => {
    expect(states).toHaveLength(6);
    expect(states[0]).toEqual([]);
    expect(states[1]).toEqual([
      { edge: "bottom", style: "continuous", weight: "thin", color: line },
    ]);
    expect(states[2]).toEqual([
      { edge: "bottom", style: "continuous", weight: "medium", color: line },
    ]);
    expect(states[3]).toEqual([
      { edge: "top", style: "continuous", weight: "thin", color: line },
      { edge: "bottom", style: "double", weight: "thin", color: line },
    ]);
    expect(states[4]?.map((entry) => entry.edge)).toEqual(OUTER);
    expect(states[5]?.map((entry) => entry.edge)).toEqual(ALL);
  });

  it("draws every line in the brand primary", () => {
    const custom = buildBorderCycle({
      ...DEFAULT_SETTINGS,
      primary: "#123456",
    });
    for (const state of custom) {
      for (const entry of state) expect(entry.color).toBe("#123456");
    }
  });

  it("finds every state it drew", () => {
    states.forEach((state, index) => {
      expect(matchBorderIndex(readback(state), states)).toBe(index);
    });
  });

  it("steps through all six looks and wraps back to none", () => {
    let index = 0;
    const walked: number[] = [];
    for (let press = 0; press < 6; press += 1) {
      const state = states[(index + 1) % states.length];
      index = matchBorderIndex(readback(state ?? []), states);
      walked.push(index);
    }
    expect(walked).toEqual([1, 2, 3, 4, 5, 0]);
  });

  it("tells a total rule from a plain one by weight alone", () => {
    const thin = readback(states[1] ?? []);
    const medium = readback(states[2] ?? []);
    expect(matchBorderIndex(thin, states)).toBe(1);
    expect(matchBorderIndex(medium, states)).toBe(2);
  });

  it("counts a partial or foreign border as the empty state", () => {
    const halfBox: BorderReadouts = {
      ...readback([]),
      top: { style: "Continuous", weight: "Thin", color: line },
      left: { style: "Continuous", weight: "Thin", color: line },
    };
    expect(matchBorderIndex(halfBox, states)).toBe(0);

    const foreign = readback(states[1] ?? []);
    foreign.bottom = { style: "Dash", weight: "Thin", color: line };
    expect(matchBorderIndex(foreign, states)).toBe(0);

    const otherColor = readback(states[1] ?? []);
    otherColor.bottom = {
      style: "Continuous",
      weight: "Thin",
      color: "#FF0000",
    };
    expect(matchBorderIndex(otherColor, states)).toBe(0);

    // A range whose cells disagree reports no style at all.
    const mixed = readback(states[4] ?? []);
    mixed.top = { style: "", weight: "", color: "" };
    expect(matchBorderIndex(mixed, states)).toBe(0);
  });

  it("reads Excel's own spelling of a line", () => {
    const shouted = readback(states[1] ?? []);
    shouted.bottom = { style: "CONTINUOUS", weight: " thin ", color: line };
    expect(matchBorderIndex(shouted, states)).toBe(1);
  });

  it("comes home from a box on a single cell, which has no inside lines", () => {
    // Nothing between the cells to report, so a box and a grid look the same:
    // the grid wins and the next press clears instead of redrawing the box.
    const box = readback(states[4] ?? [], OUTER);
    expect(matchBorderIndex(box, states)).toBe(5);
  });
});

describe("size cycles", () => {
  const { rowHeight, columnWidth } = buildSizeCycles();

  // Points, the unit Office.js takes. 15 pt is Excel's default row height; a
  // default column is 48 pt, which is the ladder's last rung, so a column steps
  // 48 -> 64 and wraps home.
  it("starts the row ladder on the Excel default", () => {
    expect(rowHeight).toEqual([15, 18, 21, 24, 30]);
    expect(columnWidth).toEqual([64, 80, 96, 120, 48]);
  });

  function walk(cycle: number[]): number[] {
    let size = cycle[0] ?? 0;
    return cycle.map(() => {
      size = nextSize(size, cycle);
      return size;
    });
  }

  it("steps each ladder through every rung and wraps home", () => {
    expect(walk(rowHeight)).toEqual([18, 21, 24, 30, 15]);
    // Columns widen for labels, then end narrow on the spacer width.
    expect(walk(columnWidth)).toEqual([80, 96, 120, 48, 64]);
  });

  it("recognises a height Excel rounded to the screen pixel", () => {
    expect(nextSize(20.75, rowHeight)).toBe(24);
    expect(nextSize(21.4, rowHeight)).toBe(24);
    // Half a point is the whole allowance: 20.4 is nobody's 21.
    expect(nextSize(20.4, rowHeight)).toBe(15);
    expect(nextSize(20.4, rowHeight, 1)).toBe(24);
  });

  it("starts a hand-dragged size at the foot of the ladder", () => {
    expect(nextSize(9, rowHeight)).toBe(15);
    expect(nextSize(255, rowHeight)).toBe(15);
    expect(nextSize(34, columnWidth)).toBe(64);
    // Nothing to step to leaves the size exactly as it is.
    expect(nextSize(18, [])).toBe(18);
  });
});

describe("hygiene cycles", () => {
  it("lists the three ladders the hygiene buttons step", () => {
    expect(INDENT_CYCLE).toEqual([0, 1, 2, 3]);
    expect(ALIGN_CYCLE).toEqual(["Left", "Center", "Right", "General"]);
    expect(UNDERLINE_CYCLE).toEqual(["Single", "Double", "None"]);
  });

  it("steps the indent one level at a time and wraps home", () => {
    expect([0, 1, 2, 3].map((level) => nextIndent(level))).toEqual([
      1, 2, 3, 0,
    ]);
  });

  it("reads a null indent, and one nobody set, as no indent", () => {
    expect(nextIndent(null)).toBe(1);
    expect(nextIndent(undefined)).toBe(1);
    // A hand-set level outside the ladder steps to its foot, the way the other
    // cycles treat a look we did not apply.
    expect(nextIndent(7)).toBe(0);
    expect(nextIndent(-2)).toBe(1);
  });

  it("steps the alignment Left, Center, Right, then back to General", () => {
    expect(
      ["Left", "Center", "Right", "General"].map((at) => nextAlignment(at)),
    ).toEqual(["Center", "Right", "General", "Left"]);
  });

  it("compares the alignment Excel gives back without minding its casing", () => {
    expect(nextAlignment("left")).toBe("Center");
    expect(nextAlignment(" CENTER ")).toBe("Right");
    // A range whose cells disagree reports nothing, and an alignment we do not
    // cycle (Fill, Justify) is nobody's rung: both start at Left.
    expect(nextAlignment(null)).toBe("Left");
    expect(nextAlignment("Justify")).toBe("Left");
  });

  it("steps the underline Single, Double, then off", () => {
    expect(["Single", "Double", "None"].map((at) => nextUnderline(at))).toEqual(
      ["Double", "None", "Single"],
    );
  });

  it("counts an accounting underline as the plain one it draws", () => {
    // Excel answers SingleAccountant for the accounting underline; it is the
    // same rung of the cycle (lessons 2026-08-27).
    expect(nextUnderline("SingleAccountant")).toBe("Double");
    expect(nextUnderline("DoubleAccountant")).toBe("None");
    expect(nextUnderline(null)).toBe("Single");
    expect(canonicalUnderline("SingleAccountant")).toBe("single");
    expect(canonicalAlignment(" Center ")).toBe("center");
  });
});
