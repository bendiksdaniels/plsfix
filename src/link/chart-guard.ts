// The chart field of a picture payload: optional cargo, never a reason to
// throw the picture away. Owns the validation of that one field on the way
// out of decodePayload, so model.ts's picture validator never sees it.
// Invariant: what comes back is a picture with a chart the slide can draw, or
// one with no chart and a chartIssue saying so - the payload's own wins.

import { isChartData } from "./chart-model";

// The reason a chart that travelled with the picture is not drawn: the panes
// print it after "as a picture", the way every other chartIssue is printed.
export const CHART_UNREADABLE = "chart data unreadable";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

// Runs on the parsed JSON before the payload validator, because a chart the
// validator refuses would otherwise take the whole payload - picture and all -
// down with it. Anything but a picture carrying a broken chart is handed back
// untouched, the same object, so a good payload costs nothing.
export function guardChart(value: unknown): unknown {
  if (!isRecord(value) || value.kind !== "picture") return value;
  if (value.chart === undefined || isChartData(value.chart)) return value;
  const guarded: Record<string, unknown> = { ...value };
  delete guarded.chart;
  // The payload's own reason wins: Excel knows better than we do why the
  // chart it sent cannot be drawn.
  if (!("chartIssue" in guarded)) guarded.chartIssue = CHART_UNREADABLE;
  return guarded;
}
