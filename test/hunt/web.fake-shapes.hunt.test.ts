// The fake host's opt-in web mode against what the rig read back in Excel for
// the web on 27.09 (.superpowers/rig/webshapes-20260926T225726.json and
// webborders-20260926T230053.json): the shapes and write answers below are
// copied from those files, so the fake can only drift from them loudly.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";

enableStrictLoadSemantics();

let helpers: FakeHelpers;

const T = "Microsoft.ExcelServices.";
const FILL = { color: true, pattern: true, patternColor: true };

beforeEach(() => {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Probe"], web: true }).helpers;
  helpers.setFill("Probe!A1", { color: "#FFFF00", pattern: "Solid" });
  helpers.setFill("Probe!A3", {
    color: "#FFFFFF",
    pattern: "LightHorizontal",
    patternColor: "#0057B8",
  });
});

async function read<T>(work: (range: Excel.Range) => T): Promise<T> {
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets
      .getItem("Probe")
      .getRange("A1:A3");
    const result = work(range);
    await context.sync();
    return result;
  });
}

describe("the web fake reads back what the rig read", () => {
  it("answers getCellProperties fills the way the web does", async () => {
    const cells = await read((range) =>
      range.getCellProperties({ format: { fill: FILL } }),
    );
    const fills = cells.value.map((row) => row[0]?.format?.fill);
    expect(fills).toEqual([
      {
        "@odata.type": `${T}CellPropertiesFill`,
        color: "#FFFF00",
        pattern: null,
        patternColor: "",
      },
      {
        "@odata.type": `${T}CellPropertiesFill`,
        color: "",
        pattern: null,
        patternColor: "",
      },
      {
        "@odata.type": `${T}CellPropertiesFill`,
        color: "#FFFFFF",
        pattern: "LightHorizontal",
        patternColor: "#0057B8",
      },
    ]);
  });

  it("annotates the border collection and answers eight edges", async () => {
    const cells = await read((range) =>
      range.getCellProperties({
        format: { borders: { color: true, style: true, weight: true } },
      }),
    );
    const borders = cells.value[1]?.[0]?.format?.borders as Record<
      string,
      unknown
    >;
    expect(borders["@odata.type"]).toBe(`${T}CellBorderCollection`);
    const edges = Object.keys(borders).filter((key) => !key.startsWith("@"));
    expect(edges.sort()).toEqual(
      [
        "bottom",
        "diagonalDown",
        "diagonalUp",
        "horizontal",
        "left",
        "right",
        "top",
        "vertical",
      ].sort(),
    );
    expect(borders.top).toEqual({
      "@odata.type": `${T}CellBorder`,
      color: "#000000",
      style: "None",
      weight: "Thin",
    });
  });

  it("answers format.fill with null where the cell answers an empty string", async () => {
    const fills = await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probe");
      const cells = ["A1", "A2"].map((address) => sheet.getRange(address));
      for (const cell of cells) {
        cell.load(
          "format/fill/color,format/fill/pattern,format/fill/patternColor",
        );
      }
      await context.sync();
      return cells.map(({ format: { fill } }) => ({
        color: fill.color,
        pattern: fill.pattern,
        patternColor: fill.patternColor,
      }));
    });
    expect(fills).toEqual([
      { color: "#FFFF00", pattern: null, patternColor: null },
      { color: "", pattern: null, patternColor: null },
    ]);
  });
});

describe("the web fake takes writes the way the rig saw them land", () => {
  async function write(address: string, props: object): Promise<string> {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Probe");
        sheet
          .getRange(address)
          .setCellProperties([[props as Excel.SettableCellProperties]]);
        await context.sync();
      });
      return "ok";
    } catch (error) {
      return (error as { code?: string }).code ?? "?";
    }
  }

  it("refuses a borders collection that carries an annotation (D6)", async () => {
    const answer = await write("D6", {
      format: { borders: { "@odata.type": {}, top: { style: "Continuous" } } },
    });
    expect(answer).toBe("InvalidArgument");
  });

  it("leaves a solid white fill for a white None pattern, clears for None alone (D1, D2)", async () => {
    helpers.setFill("Probe!D1:D2", { color: "#FFFF00", pattern: "Solid" });
    expect(await write("D1", { format: { fill: { pattern: "None" } } })).toBe(
      "ok",
    );
    const white = {
      color: "#FFFFFF",
      pattern: "None",
      patternColor: "#FFFFFF",
    };
    expect(await write("D2", { format: { fill: white } })).toBe("ok");
    expect(helpers.fill("Probe!D1").pattern).toBe("None");
    expect(helpers.fill("Probe!D2")).toMatchObject({
      color: "#FFFFFF",
      pattern: "Solid",
    });
  });

  it("draws a None edge sent with a weight, and not one sent alone (B2, B4, B6)", async () => {
    const plain = { style: "None", weight: "Thin", color: "#000000" };
    await write("B2", {
      format: { borders: { top: plain, diagonalDown: plain } },
    });
    expect(helpers.border("Probe!B2", "top").style).toBe("Continuous");
    expect(helpers.border("Probe!B2", "diagonalDown").style).toBe("None");

    helpers.sheet("Probe").edit(5, 1).borders.bottom = {
      style: "Continuous",
      weight: "Medium",
      color: "#C00000",
    };
    await write("B6", { format: { borders: { bottom: { style: "None" } } } });
    expect(helpers.border("Probe!B6", "bottom")).toEqual({
      style: "None",
      weight: "Thin",
      color: "#000000",
    });
  });
});

describe("the spill helper", () => {
  it("answers a blank spill cell as String and a truly empty one as Empty", async () => {
    helpers.spill("Probe!C1", '=IF(ROW(1:3)>2,ROW(1:3),"")', [[""], [""], [3]]);
    const types = await Excel.run(async (context) => {
      const range = context.workbook.worksheets
        .getItem("Probe")
        .getRange("C1:C4");
      range.load("values,formulas,valueTypes");
      await context.sync();
      return {
        values: range.values,
        formulas: range.formulas,
        types: range.valueTypes,
      };
    });
    expect(types.values).toEqual([[""], [""], [3], [""]]);
    expect(types.formulas).toEqual([
      ['=IF(ROW(1:3)>2,ROW(1:3),"")'],
      [""],
      [""],
      [""],
    ]);
    expect(types.types).toEqual([
      ["String"],
      ["String"],
      ["Integer"],
      ["Empty"],
    ]);
  });
});
