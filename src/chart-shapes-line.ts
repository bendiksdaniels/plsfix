// The native line chart's primitives: one small ellipse marker per point, a
// connector between consecutive points, and the point's own value centred
// above its marker (or below it where the marker sits on the plot's top
// edge). Pure: no Office.js, no DOM. Invariant: every primitive lies inside
// the box it was given.

import { boxAt, label, lineBetween } from "./chart-shapes-parts";
import {
  LABEL_HEIGHT,
  LABEL_SIZE,
  textWidth,
  valueY,
  type Primitive,
  type ValueScale,
} from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData, ChartSeries } from "./link/chart-model";

const MARKER_MIN = 3;
const MARKER_MAX = 6;
const MARKER_SPACING_DIVISOR = 5;
const CONNECTOR_WEIGHT = 1.5;
// How close to the plot's own top a marker has to sit before its label would
// spill out above the box and has to flip below the marker instead.
const TOP_EDGE_EPSILON = 0.01;

interface Point {
  left: number;
  top: number;
}

// Point i sits at the centre of slot i, the same centring categoryLabels in
// chart-shapes.ts already uses for the axis below it (Excel's own default
// for a category axis): half a slot in from the edge, not flush with it.
function positionsFor(
  series: ChartSeries,
  slot: number,
  plot: Box,
  scale: ValueScale,
): Point[] {
  return series.values.map((value, pointIndex) => ({
    left: plot.left + (pointIndex + 0.5) * slot,
    top: valueY(value, scale, plot),
  }));
}

// A marker never runs past the plot: its centre follows the point, but the
// box itself is pulled back to the nearest edge once it would spill out.
function markerBoxAt(point: Point, marker: number, plot: Box): Box {
  return boxAt(
    Math.max(
      plot.left,
      Math.min(point.left - marker / 2, plot.left + plot.width - marker),
    ),
    Math.max(
      plot.top,
      Math.min(point.top - marker / 2, plot.top + plot.height - marker),
    ),
    marker,
    marker,
  );
}

// The point's value, centred on its marker: above it normally, below it when
// the marker is already flush with the plot's top edge.
function pointLabel(
  content: string,
  ink: string,
  point: Point,
  markerBox: Box,
  plot: Box,
  name: string,
): Primitive {
  const width = textWidth(content, LABEL_SIZE);
  const left = Math.max(
    plot.left,
    Math.min(point.left - width / 2, plot.left + plot.width - width),
  );
  const touchesTop = markerBox.top <= plot.top + TOP_EDGE_EPSILON;
  const top = touchesTop
    ? markerBox.top + markerBox.height
    : markerBox.top - LABEL_HEIGHT;
  return label(boxAt(left, top, width, LABEL_HEIGHT), content, ink, "c", name);
}

// One series' connectors, markers and value labels. A connector joins a point
// to the one before it, so the first point in the series draws none.
function seriesPrimitives(
  series: ChartSeries,
  seriesIndex: number,
  ink: string,
  plot: Box,
  scale: ValueScale,
  slot: number,
  marker: number,
): Primitive[] {
  const positions = positionsFor(series, slot, plot, scale);
  const out: Primitive[] = [];
  positions.forEach((point, pointIndex) => {
    if (pointIndex > 0) {
      const previous = positions[pointIndex - 1]!;
      out.push(
        lineBetween(
          previous.left,
          previous.top,
          point.left,
          point.top,
          series.colors[pointIndex]!,
          CONNECTOR_WEIGHT,
          `line ${seriesIndex}.${pointIndex - 1}`,
        ),
      );
    }
    const markerBox = markerBoxAt(point, marker, plot);
    out.push({
      kind: "ellipse",
      box: markerBox,
      color: series.colors[pointIndex]!,
      name: `marker ${seriesIndex}.${pointIndex}`,
    });
    out.push(
      pointLabel(
        series.labels[pointIndex]!,
        ink,
        point,
        markerBox,
        plot,
        `label ${seriesIndex}.${pointIndex}`,
      ),
    );
  });
  return out;
}

// A native line chart becomes one editable connector and one small marker per
// point, each with its own value label. Markers make individual values
// selectable in PowerPoint and avoid a line disappearing into a
// matching-colour slide background.
export function lineSeries(
  data: ChartData,
  plot: Box,
  scale: ValueScale,
): Primitive[] {
  const points = data.categories.length;
  const slot = plot.width / points;
  const marker = Math.min(
    MARKER_MAX,
    Math.max(MARKER_MIN, slot / MARKER_SPACING_DIVISOR),
  );
  const out: Primitive[] = [];
  data.series.forEach((series, seriesIndex) => {
    out.push(
      ...seriesPrimitives(
        series,
        seriesIndex,
        data.ink,
        plot,
        scale,
        slot,
        marker,
      ),
    );
  });
  return out;
}
