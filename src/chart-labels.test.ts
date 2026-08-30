// Where a chart's value labels sit, decided by chart type: outside the end of
// a clustered bar or a pie slice, inside a stacked segment, above a line point,
// and nowhere at all where Excel offers no position (waterfall, doughnut, area).

import { describe, expect, it } from "vitest";
import { labelPosition, leaderLines } from "./chart-labels";

describe("leaderLines", () => {
  it("ties pie labels to their slices and nothing else", () => {
    expect(leaderLines("Pie")).toBe(true);
    expect(leaderLines("3DPieExploded")).toBe(true);
    expect(leaderLines("Doughnut")).toBe(false);
    expect(leaderLines("ColumnClustered")).toBe(false);
  });
});

describe("labelPosition", () => {
  it("puts clustered column and bar labels outside the end", () => {
    for (const type of [
      "ColumnClustered",
      "BarClustered",
      "3DColumnClustered",
      "CylinderColClustered",
      "ConeBarClustered",
    ]) {
      expect(labelPosition(type), type).toBe("OutsideEnd");
    }
  });

  it("puts pie labels outside the slice", () => {
    for (const type of [
      "Pie",
      "PieExploded",
      "3DPie",
      "3DPieExploded",
      "PieOfPie",
      "BarOfPie",
    ]) {
      expect(labelPosition(type), type).toBe("OutsideEnd");
    }
  });

  it("centres a label inside a stacked segment", () => {
    for (const type of [
      "ColumnStacked",
      "ColumnStacked100",
      "BarStacked",
      "3DBarStacked100",
      "PyramidColStacked",
    ]) {
      expect(labelPosition(type), type).toBe("Center");
    }
  });

  it("lifts line, scatter and radar labels above their points", () => {
    for (const type of [
      "Line",
      "LineMarkers",
      "LineStacked",
      "XYScatter",
      "XYScatterLinesNoMarkers",
      "Radar",
      "RadarMarkers",
    ]) {
      expect(labelPosition(type), type).toBe("Top");
    }
  });

  it("leaves the position alone where Excel offers none", () => {
    for (const type of [
      "Waterfall",
      "Doughnut",
      "DoughnutExploded",
      "Area",
      "AreaStacked",
      "3DColumn",
      "CylinderCol",
      "Bubble",
      "Surface",
      "StockHLC",
      "Treemap",
      "Sunburst",
      "Funnel",
      "Histogram",
      "Invalid",
      "",
    ]) {
      expect(labelPosition(type), type).toBeNull();
    }
  });
});
