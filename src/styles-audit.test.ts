import { describe, expect, it } from "vitest";
import { unusedStyles, type WorkbookStyle } from "./styles-audit";

const table: WorkbookStyle[] = [
  { name: "Normal", builtIn: true },
  { name: "Comma", builtIn: true },
  { name: "Input box", builtIn: false },
  { name: "Assumption", builtIn: false },
];

describe("unused styles", () => {
  it("never offers a built-in style, however unused", () => {
    expect(unusedStyles(table, new Set())).toEqual(["Assumption", "Input box"]);
  });

  it("leaves out a style a cell is wearing", () => {
    expect(unusedStyles(table, new Set(["Assumption"]))).toEqual(["Input box"]);
  });

  it("counts a style as used whatever case the cell reports", () => {
    expect(unusedStyles(table, new Set(["input BOX"]))).toEqual(["Assumption"]);
  });

  it("sorts the names so the list reads the same on every scan", () => {
    const shuffled: WorkbookStyle[] = [
      { name: "zebra", builtIn: false },
      { name: "Alpha", builtIn: false },
      { name: "beta", builtIn: false },
    ];
    expect(unusedStyles(shuffled, new Set())).toEqual([
      "Alpha",
      "beta",
      "zebra",
    ]);
  });

  it("finds nothing to clean in a workbook of built-ins only", () => {
    const builtIns = table.filter((style) => style.builtIn);
    expect(unusedStyles(builtIns, new Set(["Normal"]))).toEqual([]);
  });
});
