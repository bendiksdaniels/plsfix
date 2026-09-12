// Direct proof of the six calculation-block grids in template-blocks.ts:
// shape, placeholder bounds, number formats and text/formula typing that
// src/excel/templates.ts relies on to write a block onto a sheet, plus the
// EBITDA bridge's fit for the two-column shape insertWaterfall
// (src/excel/charts.ts) selects off it.
import { describe, expect, it } from "vitest";
import { placeholderRefs } from "./template-cells";
import { TEMPLATES } from "./templates";

// insertWaterfall accepts two adjacent columns, labels left and values right,
// with at least this many rows (src/excel/charts.ts's own BRIDGE_MIN_POINTS).
const WATERFALL_MIN_POINTS = 3;

// Kinds a cell holds only as a name or a result label, never as a number:
// "header" is left out because the DCF and NPV blocks number their own
// columns with numeric header cells (a forecast year, a period index).
const TEXTUAL_KINDS = new Set(["title", "label", "result"]);

describe.each(TEMPLATES.map((template) => [template.id, template] as const))(
  "%s",
  (_id, template) => {
    it("is a rectangle of the size it declares", () => {
      expect(template.cells).toHaveLength(template.rows);
      for (const row of template.cells) {
        expect(row).toHaveLength(template.cols);
      }
    });

    it("keeps every placeholder inside the block and every formula an assignment", () => {
      template.cells.forEach((row, rowIndex) => {
        row.forEach((cell, columnIndex) => {
          if (cell.f === undefined) return;
          const where = `${template.id} r${String(rowIndex)}c${String(columnIndex)}: ${cell.f}`;
          expect(cell.f.startsWith("="), where).toBe(true);
          for (const ref of placeholderRefs(cell.f)) {
            expect(ref.row, where).toBeGreaterThanOrEqual(0);
            expect(ref.row, where).toBeLessThan(template.rows);
            expect(ref.column, where).toBeGreaterThanOrEqual(0);
            expect(ref.column, where).toBeLessThan(template.cols);
          }
        });
      });
    });

    it("gives every numeric cell - a literal or a formula - a non-empty number format", () => {
      for (const row of template.cells) {
        for (const cell of row) {
          const numeric = typeof cell.v === "number" || cell.f !== undefined;
          if (!numeric) continue;
          expect(typeof cell.format).toBe("string");
          expect(cell.format).not.toBe("");
        }
      }
    });

    it("keeps every title, label or result cell that carries a value as text", () => {
      for (const row of template.cells) {
        for (const cell of row) {
          if (!TEXTUAL_KINDS.has(cell.kind) || cell.f !== undefined) continue;
          if (cell.v === undefined) continue;
          expect(typeof cell.v).toBe("string");
        }
      }
    });
  },
);

describe("the EBITDA bridge", () => {
  const bridge = TEMPLATES.find((template) => template.id === "ebitda-bridge")!;
  // Row 0 is the title and the last row is the reconciling "Check" formula:
  // neither belongs to the range a modeller drags over the bridge before
  // charting it, so the shape that matters is everything between them.
  const span = bridge.cells.slice(1, bridge.rows - 1);

  it("is two columns wide, labels left and values right, the shape insertWaterfall reads", () => {
    expect(bridge.cols).toBe(2);
    for (const row of span) {
      expect(typeof row[0]?.v).toBe("string");
      expect(typeof row[1]?.v).toBe("number");
    }
  });

  it("has enough rows for insertWaterfall's bridge shape once the title and check row are excluded", () => {
    expect(span.length).toBeGreaterThanOrEqual(WATERFALL_MIN_POINTS);
  });
});
