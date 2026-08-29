// The templates contract: every block is rectangular, every formula names a
// cell inside its own block, and the same placeholder grid lands wherever the
// modeller happens to stand. The samples pin one formula per template to the
// A1 string it becomes at C5, so a renumbered row cannot pass unnoticed.
import { describe, expect, it } from "vitest";
import {
  placeholderRefs,
  resolveFormula,
  type Template,
  templateById,
  TEMPLATES,
} from "./templates";

// Column index 2, row index 4: the origin every sample below resolves against.
const C5 = { row: 4, column: 2 };

function formulaAt(template: Template, row: number, column: number): string {
  const cell = template.cells[row]?.[column];
  if (!cell?.f) {
    throw new Error(`${template.id} has no formula at ${row},${column}`);
  }
  return cell.f;
}

function resolvedAt(
  id: string,
  row: number,
  column: number,
  origin = C5,
): string {
  const template = templateById(id);
  if (!template) throw new Error(`no template ${id}`);
  return resolveFormula(formulaAt(template, row, column), origin);
}

describe("resolveFormula", () => {
  it("reads an offset from the block's own origin", () => {
    expect(resolveFormula("={r+1,c}", C5)).toBe("=C6");
    expect(resolveFormula("={r,c+3}", C5)).toBe("=F5");
  });

  it("moves every reference when the block moves", () => {
    expect(resolveFormula("={r+1,c}", { row: 0, column: 0 })).toBe("=A2");
    expect(resolveFormula("={r+1,c}", { row: 9, column: 26 })).toBe("=AA11");
    expect(resolveFormula("={r+1,c+1}", { row: 99, column: 27 })).toBe(
      "=AC101",
    );
  });

  it("resolves both ends of a range placeholder", () => {
    expect(resolveFormula("=SUM({r,c}:{r+2,c})", C5)).toBe("=SUM(C5:C7)");
  });

  it("leaves text that carries no placeholder alone", () => {
    expect(resolveFormula('=IFERROR(1/0,"n/a")', C5)).toBe(
      '=IFERROR(1/0,"n/a")',
    );
  });
});

describe("placeholderRefs", () => {
  it("lists every cell a formula names, in order", () => {
    expect(placeholderRefs("=SUM({r+1,c}:{r+3,c})*{r,c+2}")).toEqual([
      { row: 1, column: 0 },
      { row: 3, column: 0 },
      { row: 0, column: 2 },
    ]);
  });

  it("finds nothing in a formula without placeholders", () => {
    expect(placeholderRefs("=365")).toEqual([]);
  });
});

describe("the template catalogue", () => {
  it("holds the six blocks under unique ids", () => {
    expect(TEMPLATES).toHaveLength(6);
    expect(new Set(TEMPLATES.map((one) => one.id)).size).toBe(6);
  });

  it("answers by id and refuses an unknown one", () => {
    expect(templateById("dcf")?.name).toBe("DCF valuation");
    expect(templateById("nope")).toBeNull();
  });

  for (const template of TEMPLATES) {
    it(`${template.id} is a rectangle of the size it declares`, () => {
      expect(template.name).not.toBe("");
      expect(template.description).not.toBe("");
      expect(template.cells).toHaveLength(template.rows);
      for (const row of template.cells) {
        expect(row).toHaveLength(template.cols);
      }
    });

    it(`${template.id} references only cells inside the block`, () => {
      let formulas = 0;
      template.cells.forEach((row, rowIndex) => {
        row.forEach((cell, columnIndex) => {
          if (!cell.f) return;
          formulas += 1;
          const where = `${template.id} ${rowIndex},${columnIndex}: ${cell.f}`;
          for (const ref of placeholderRefs(cell.f)) {
            expect(ref.row, where).toBeGreaterThanOrEqual(0);
            expect(ref.row, where).toBeLessThan(template.rows);
            expect(ref.column, where).toBeGreaterThanOrEqual(0);
            expect(ref.column, where).toBeLessThan(template.cols);
          }
        });
      });
      expect(formulas).toBeGreaterThan(0);
    });

    it(`${template.id} carries a value or a formula in no cell twice`, () => {
      for (const row of template.cells) {
        for (const cell of row) {
          expect(cell.v === undefined || cell.f === undefined).toBe(true);
        }
      }
    });
  }
});

describe("a sample formula per template at C5", () => {
  it("prices the annuity payment off the four debt inputs", () => {
    expect(resolvedAt("debt-schedule", 6, 1)).toBe("=PMT(D7/D9,D10,-D6)");
  });

  it("grows the last forecast year into a Gordon terminal value", () => {
    expect(resolvedAt("dcf", 10, 1)).toBe("=H10*(1+D7)/(D6-D7)");
  });

  it("discounts the eight cash flows and adds the outlay back", () => {
    expect(resolvedAt("npv-irr", 14, 1)).toBe("=NPV(D6,D10:D17)+D9");
  });

  it("turns receivables into days of revenue", () => {
    expect(resolvedAt("working-capital", 8, 1)).toBe("=D8/D6*365");
  });

  it("steps the base value by both drivers in the grid corner", () => {
    expect(resolvedAt("sensitivity", 6, 1)).toBe("=D6*(1+C11)*(1+D10)");
  });

  it("checks the bridge back to the closing total", () => {
    expect(resolvedAt("ebitda-bridge", 8, 1)).toBe("=D6+SUM(D7:D11)-D12");
  });
});
