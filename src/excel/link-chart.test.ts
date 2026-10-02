// The pure halves of the chart reader: which Excel chart types a slide can
// draw as shapes, and what a dimension's raw cells are worth as numbers.
// Everything else in link-chart.ts talks to office.js and is covered by
// test/links.chart.integration.test.ts against the strict fake host.

import { describe, expect, it } from "vitest";
import { chartKind, valuesOf } from "./link-chart";

describe("chartKind", () => {
  it("maps the column families, clustered and stacked", () => {
    expect(chartKind("ColumnClustered", 0)).toBe("column");
    expect(chartKind("3DColumnClustered", null)).toBe("column");
    expect(chartKind("CylinderColClustered", null)).toBe("column");
    expect(chartKind("ColumnStacked100", null)).toBe("stackedColumn");
    expect(chartKind("PyramidColStacked", null)).toBe("stackedColumn");
  });

  it("maps the bar families before the clustered catch-all", () => {
    expect(chartKind("BarClustered", 100)).toBe("bar");
    expect(chartKind("ConeBarClustered", 0)).toBe("bar");
    expect(chartKind("BarStacked100", null)).toBe("stackedBar");
    expect(chartKind("3DBarStacked", null)).toBe("stackedBar");
  });

  it("maps the waterfall and every pie", () => {
    expect(chartKind("Waterfall", null)).toBe("waterfall");
    expect(chartKind("Pie", null)).toBe("pie");
    expect(chartKind("3DPieExploded", null)).toBe("pie");
  });

  it("leaves every other chart type to the picture", () => {
    expect(chartKind("Line", null)).toBe("line");
    expect(chartKind("LineMarkers", null)).toBe("line");
    expect(chartKind("Doughnut", null)).toBeNull();
    expect(chartKind("Area", null)).toBeNull();
    expect(chartKind("", null)).toBeNull();
  });

  // Columns are laid out side by side; a full overlap would hide one behind
  // the other, so that chart is one the picture says better.
  it("refuses columns drawn on top of each other", () => {
    expect(chartKind("ColumnClustered", 100)).toBeNull();
  });
});

describe("valuesOf", () => {
  it("reads the numbers a dimension gives back, whatever their type", () => {
    expect(valuesOf(["1", 2.5])).toEqual([1, 2.5]);
    expect(valuesOf([])).toEqual([]);
  });

  it("refuses a dimension that is not all numbers", () => {
    expect(valuesOf(["x"])).toBeNull();
    // A blank is a gap in the data, not a zero.
    expect(valuesOf(["1", "  "])).toBeNull();
  });
});
