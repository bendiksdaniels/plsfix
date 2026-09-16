// PowerPoint.Table.styleSettings (PowerPointApi 1.9): the header band, the
// defensive fix-up for a table PowerPoint handed over with no style, and
// which of a repaint's cells are still pls,fix's own to clear. Owns every
// styleSettings read and write. Invariant: a host below 1.9 never has
// styleSettings addressed at all, not even to leave it alone.

import { TAG_PAINT, type TablePayload } from "../link/model";
import { decodePaintMap, paintKey } from "../link/paint-map";
import { withSyncDeadline } from "./chart-draw";
import { hasPowerPointApi } from "./shapes";

export const TABLE_STYLE_API = "1.9";
const READING_STYLE = "reading the table style";

export function hasTableStyle(): boolean {
  return hasPowerPointApi(TABLE_STYLE_API);
}

// Both of PowerPoint's "no style" presets: a table left in either reads with
// a header band and no banding to set it off from.
function isNoStyle(style: string): boolean {
  const noStyle: string[] = [
    PowerPoint.TableStyle.noStyleNoGrid,
    PowerPoint.TableStyle.noStyleTableGrid,
  ];
  return noStyle.includes(style);
}

// A table with no style at all gets PowerPoint's own default for one added
// through the ribbon, Medium Style 2 Accent 1; any other style, including one
// the user picked, is left exactly as it is.
export function queueDefaultStyle(
  settings: PowerPoint.TableStyleSettings,
): void {
  if (isNoStyle(settings.style)) {
    settings.style = PowerPoint.TableStyle.mediumStyle2Accent1;
  }
}

// Queued, not written: the caller's next sync carries it, the same way every
// other cell and format write on a table link does. A repaint calls this
// blind, with no read first; insert and rebuild call it after readTableStyle.
export function queueHeaderRow(table: PowerPoint.Table, header: boolean): void {
  if (!hasTableStyle()) return;
  table.styleSettings.isFirstRowHighlighted = header;
}

// The one extra round trip insert and rebuild spend on a just-created table's
// own style, before either queues a fix-up on it.
export async function readTableStyle(
  context: PowerPoint.RequestContext,
  table: PowerPoint.Table,
  what: string,
): Promise<PowerPoint.TableStyleSettings> {
  const settings = table.styleSettings;
  settings.load("style");
  await withSyncDeadline(context.sync(), what);
  return settings;
}

// The style a just-created table starts with, on the API that carries one: a
// fix-up for a host that hands back no style at all, and the header band the
// payload's own row 0 asked for. Skipped whole - at no extra round trip - on
// a host below TABLE_STYLE_API, since hasTableStyle is checked before
// readTableStyle ever runs.
export async function styleNewTable(
  context: PowerPoint.RequestContext,
  table: PowerPoint.Table,
  payload: TablePayload,
): Promise<void> {
  if (!hasTableStyle()) return;
  const settings = await readTableStyle(context, table, READING_STYLE);
  queueDefaultStyle(settings);
  queueHeaderRow(table, payload.h === true);
}

// TAG_PAINT absent or undecodable reads as null: a deck painted before this
// shipped, which fillsToClear treats as today's behaviour minus the header
// band a header row now keeps.
export function readPaintTag(shape: PowerPoint.Shape): Set<string> | null {
  const value = shape.tags.items.find((item) => item.key === TAG_PAINT)?.value;
  return decodePaintMap(value);
}

// Which cells a repaint is allowed to blank. With a paint map from a prior
// push, only a cell pls,fix filled before and the source has since dropped
// the fill from. Without one - a deck painted before this shipped - every
// cell the payload leaves unfilled, except row 0 when it is a header: today's
// behaviour, minus the band a header row now keeps.
export function fillsToClear(
  previous: Set<string> | null,
  payload: TablePayload,
): Set<string> {
  const clear = new Set<string>();
  payload.cells.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (cell.f !== undefined) return;
      const key = paintKey(r, c);
      const wasPainted =
        previous === null
          ? !(r === 0 && payload.h === true)
          : previous.has(key);
      if (wasPainted) clear.add(key);
    });
  });
  return clear;
}
