// What a link's source becomes on the wire: the base64 picture a range or a
// chart gives, the cell grid link-table.ts reads, or the one cell a text link
// shows. Owns the Render union and every render. Invariant: a render bound to
// a fresh anchor undoes that anchor when it fails, so no hidden name outlives
// the export that made it.

import { type ChartData } from "../link/chart-model";
import {
  staged,
  type ResolvedChart,
  type ResolvedSource,
} from "./link-anchors";
import { readChartData } from "./link-chart";
import { isMacExcel } from "./link-platform";
import { renderTable, type TableRender } from "./link-table";

// Charts are laid out in points; rendering at twice that keeps the slide
// picture sharp on a high-density screen.
const CHART_PIXEL_SCALE = 2;

export type Render =
  | { kind: "picture"; png: string; chart?: ChartData }
  | { kind: "text"; text: string }
  | ({ kind: "table" } & TableRender);

// The anchor is bound in the same batch the picture is asked for, and office.js
// batches are not transactional: the rename or the names.add executes and only
// the getImage fails, so the workbook is already mutated when the sync rejects.
// A failed render therefore undoes its own anchor rather than leaving one no
// registry entry claims.
export async function renderAnchored(
  context: Excel.RequestContext,
  resolved: ResolvedSource,
  label: string,
  release: () => void,
): Promise<Render> {
  try {
    return await renderSource(context, resolved);
  } catch (error) {
    await undoAnchor(context, release);
    throw staged(`export ${label}`, error);
  }
}

// Best effort, like publish's rollback: if the workbook will not take the undo,
// the render error the caller is about to see is the one worth reporting.
async function undoAnchor(
  context: Excel.RequestContext,
  release: () => void,
): Promise<void> {
  try {
    release();
    await context.sync();
  } catch {
    return;
  }
}

export async function renderSource(
  context: Excel.RequestContext,
  resolved: ResolvedSource,
): Promise<Render> {
  if (resolved.kind === "table") {
    return { kind: "table", ...(await renderTable(context, resolved.range)) };
  }
  // What Excel shows in the cell, number format and separators included: the
  // slide is never asked to format a value it cannot see.
  if (resolved.kind === "text") {
    resolved.range.load("text");
    await context.sync();
    return { kind: "text", text: resolved.range.text[0]?.[0] ?? "" };
  }
  if (resolved.kind === "chart") return renderChart(context, resolved);
  const image = resolved.range.getImage();
  await context.sync();
  return { kind: "picture", png: image.value };
}

// The picture is queued first and the reads ride the same batch: a chart the
// slide cannot draw costs the one round trip a chart link costs today.
async function renderChart(
  context: Excel.RequestContext,
  resolved: ResolvedChart,
): Promise<Render> {
  // Excel Mac 16.107 has reported GeneralException for the optional sizing
  // arguments even though getImage itself is available. Its default image is
  // the chart's rendered size and is reliable; other hosts retain the sharper
  // two-times export used for slide placement.
  const image = isMacExcel()
    ? resolved.chart.getImage()
    : resolved.chart.getImage(
        Math.round(resolved.width * CHART_PIXEL_SCALE),
        Math.round(resolved.height * CHART_PIXEL_SCALE),
        Excel.ImageFittingMode.fit,
      );
  const chart = await readChartData(context, resolved.chart);
  return { kind: "picture", png: image.value, ...(chart ? { chart } : {}) };
}
