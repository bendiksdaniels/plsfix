// Where a chart's value labels sit, by chart type. Pure: an Excel chart type
// name in, a ChartDataLabelPosition name or null out; null means Excel offers
// that chart no position and the caller sets none. Owns the type-name grammar.

export type LabelPosition = "OutsideEnd" | "Center" | "Top";

// Lines, scatters and radars label above the point; every pie relative outside
// the slice; clustered columns and bars (3-D, cylinder, cone and pyramid ones
// included) outside the end; stacked ones in the middle of their segment.
// Everything else - waterfall and the other chartex charts, doughnuts, areas,
// surfaces, bubbles, stock charts - keeps the host's own placement.
export function labelPosition(chartType: string): LabelPosition | null {
  if (/^(Line|XYScatter|Radar)/.test(chartType)) return "Top";
  if (chartType.includes("Pie")) return "OutsideEnd";
  if (chartType.endsWith("Clustered")) return "OutsideEnd";
  if (
    /^(3D)?(Column|Bar|Cylinder|Cone|Pyramid)\w*Stacked(100)?$/.test(chartType)
  ) {
    return "Center";
  }
  return null;
}

// A label outside a slice needs the line that ties it back; nothing else does.
export function leaderLines(chartType: string): boolean {
  return chartType.includes("Pie");
}
