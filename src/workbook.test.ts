import { describe, expect, it } from "vitest";
import { brokenNames, tocRows } from "./workbook";

describe("tocRows", () => {
  it("numbers visible sheets from one", () => {
    expect(
      tocRows([
        { name: "Cover", visibility: "Visible" },
        { name: "Model", visibility: "Visible" },
      ]),
    ).toEqual([
      { index: 1, name: "Cover", target: "'Cover'!A1" },
      { index: 2, name: "Model", target: "'Model'!A1" },
    ]);
  });

  it("quotes names with spaces", () => {
    expect(tocRows([{ name: "P&L build", visibility: "Visible" }])).toEqual([
      { index: 1, name: "P&L build", target: "'P&L build'!A1" },
    ]);
  });

  it("doubles embedded apostrophes", () => {
    expect(tocRows([{ name: "Bank's view", visibility: "Visible" }])).toEqual([
      { index: 1, name: "Bank's view", target: "'Bank''s view'!A1" },
    ]);
  });

  it("skips hidden and very hidden sheets without spending an index", () => {
    expect(
      tocRows([
        { name: "Cover", visibility: "Visible" },
        { name: "Scratch", visibility: "Hidden" },
        { name: "Archive", visibility: "VeryHidden" },
        { name: "Model", visibility: "Visible" },
      ]),
    ).toEqual([
      { index: 1, name: "Cover", target: "'Cover'!A1" },
      { index: 2, name: "Model", target: "'Model'!A1" },
    ]);
  });

  it("returns nothing for an empty workbook", () => {
    expect(tocRows([])).toEqual([]);
  });

  it("returns nothing when every sheet is hidden", () => {
    expect(tocRows([{ name: "Scratch", visibility: "Hidden" }])).toEqual([]);
  });
});

describe("brokenNames", () => {
  it("returns names whose reference is broken", () => {
    expect(
      brokenNames([
        { name: "Sales", formula: "=Model!#REF!" },
        { name: "Margin", formula: "=Model!$B$4" },
      ]),
    ).toEqual(["Sales"]);
  });

  it("leaves healthy names alone", () => {
    expect(
      brokenNames([
        { name: "Margin", formula: "=Model!$B$4" },
        { name: "Rate", formula: "=0.05" },
        { name: "Label", formula: '="#REFERENCE"' },
      ]),
    ).toEqual([]);
  });

  it("catches a broken reference anywhere in the formula", () => {
    expect(
      brokenNames([
        { name: "Span", formula: "=SUM(Model!$B$4:#REF!)" },
        { name: "Sheet", formula: "=#REF!$A$1" },
      ]),
    ).toEqual(["Span", "Sheet"]);
  });

  it("returns nothing for a workbook with no names", () => {
    expect(brokenNames([])).toEqual([]);
  });
});
