// PowerPoint.Table.styleSettings (PowerPointApi 1.9): the header band and
// which of a repaint's cells are still pls,fix's own to clear. Owns every
// styleSettings write. Invariant: a host below 1.9 never has styleSettings
// addressed at all, not even to leave it alone; nothing here ever reads it
// back (see queueTableStyle for why).

import { TAG_PAINT, type TablePayload } from "../link/model";
import { decodePaintMap, paintKey } from "../link/paint-map";
import { hasPowerPointApi } from "./shapes";

export const TABLE_STYLE_API = "1.9";

export function hasTableStyle(): boolean {
  return hasPowerPointApi(TABLE_STYLE_API);
}

// A table shapes.addTable hands back on PowerPoint for Mac (16.107,
// PowerPointApi 1.8, 1.9 and 1.10 all isSetSupported) starts with NO style at
// all - the saved XML shows an empty <a:tblPr/> - and on such a table every
// styleSettings.load(...) is refused with GeneralException, style alone
// included, until a style has been written. Reading is therefore never an
// option: this writes PowerPoint's own ribbon default (Medium Style 2
// Accent 1) blind, then the header flag from the payload, then every other
// flag to a plain, unbanded body. Queued, not read back: the caller's next
// sync carries all seven, the same way every other cell and format write on
// a table link travels to the next sync. The look this produces is
// deterministic - the header band comes from the style, the body is plain,
// and the payload's own cell fills sit on top - whatever style, if any, the
// table started with. Skipped whole, at no round trip, on a host below
// TABLE_STYLE_API.
export function queueTableStyle(
  table: PowerPoint.Table,
  header: boolean,
): void {
  if (!hasTableStyle()) return;
  const settings = table.styleSettings;
  settings.style = PowerPoint.TableStyle.mediumStyle2Accent1;
  settings.isFirstRowHighlighted = header;
  settings.areRowsBanded = false;
  settings.areColumnsBanded = false;
  settings.isFirstColumnHighlighted = false;
  settings.isLastRowHighlighted = false;
  settings.isLastColumnHighlighted = false;
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
