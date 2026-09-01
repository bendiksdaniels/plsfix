// What a link's source becomes on the wire: the base64 picture a range or a
// chart gives, the cell grid link-table.ts reads, or the one cell a text link
// shows. Owns the Render union and every render. A call a host refuses is
// retried once, plainly, in the same context - what the host answers decides,
// never which platform it says it is. Invariant: a render bound to a fresh
// anchor undoes that anchor when every attempt fails, so no hidden name
// outlives the export that made it.

import { type ChartData } from "../link/chart-model";
import {
  bothRefused,
  staged,
  type ResolvedChart,
  type ResolvedSource,
} from "./link-anchors";
import { readChartData } from "./link-chart";
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
// slide cannot draw costs the one round trip a chart link costs today. A host
// that refuses that batch - the sizing arguments, or a head read beside them -
// is asked again, plainly, rather than costing the link its picture.
async function renderChart(
  context: Excel.RequestContext,
  resolved: ResolvedChart,
): Promise<Render> {
  try {
    const image = resolved.chart.getImage(
      Math.round(resolved.width * CHART_PIXEL_SCALE),
      Math.round(resolved.height * CHART_PIXEL_SCALE),
      Excel.ImageFittingMode.fit,
    );
    const chart = await readChartData(context, resolved.chart);
    return picture(image.value, chart);
  } catch (sharp) {
    return renderChartPlain(context, resolved, sharp);
  }
}

// The picture alone, at the chart's own size, in the context the sharp attempt
// already failed in. It gets its own sync first, so the chart data - the read
// that may have been what failed - can fail again without taking the picture
// with it, exactly as an undrawable chart travels today.
async function renderChartPlain(
  context: Excel.RequestContext,
  resolved: ResolvedChart,
  sharp: unknown,
): Promise<Render> {
  try {
    const image = resolved.chart.getImage();
    await context.sync();
    return picture(image.value, await tryChartData(context, resolved.chart));
  } catch (plain) {
    throw bothRefused(sharp, "sharp", plain, "plain");
  }
}

// readChartData answers null for a chart it cannot describe, but a host that
// refuses the read itself rejects the sync: the picture is already in hand by
// then, and a link with a picture beats one with nothing.
async function tryChartData(
  context: Excel.RequestContext,
  chart: Excel.Chart,
): Promise<ChartData | null> {
  try {
    return await readChartData(context, chart);
  } catch {
    return null;
  }
}

// One shape for both attempts: the chart key is left out entirely rather than
// sent as null, because that is what the payload codec reads as "no chart".
function picture(png: string, chart: ChartData | null): Render {
  return { kind: "picture", png, ...(chart ? { chart } : {}) };
}
