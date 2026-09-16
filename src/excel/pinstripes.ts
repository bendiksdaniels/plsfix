// Pinstripes: every second row (or column) of the selection banded in a very
// light brand tint, so a wide comps or schedule grid reads across without a
// ruler. A second press over bands that are all still exactly this tint takes
// them off again.
//
// Owns: the band tint, the on/off decision read off the fills themselves, and
// the paint. Invariant: it never paints over an overlay's fills - the audit
// stripes and the linked-cell highlight each hand back what they covered, and
// a band on top of one would be given back as the modeller's own formatting.

import { intersects } from "../link/geometry";
import {
  applyFillKey,
  hostSupports,
  selectedSingleRange,
  withinCap,
} from "./internal";
import { fillGrid, requestFills, requireNoOverlayOwner } from "./fill-store";
import { paintSync, protectedNote, sheetProtected } from "./protection";
import { parseAddress } from "./shared";
import { captureUndo } from "./undo";
import { getActiveSettings, tint } from "../settings";

/** Which way the bands run: down the rows, or across the columns. */
export type PinstripeAxis = "rows" | "columns";

const STAGE = "Pinstripes";
// 10 % of the palette's primary mixed over white - the same arithmetic as
// round(255 - (255 - c) * 0.10) per channel - which is light enough to read
// numbers over and dark enough to follow a row across a wide grid.
const BAND_TINT = 0.9;
// Rows 2, 4, 6...: the first line of the selection stays clear, so a header
// row a modeller included keeps whatever look it already has.
const FIRST_BAND = 1;
const BAND_STEP = 2;
const MIN_LINES = 2;

function bandKey(): string {
  const color = tint(getActiveSettings().primary, BAND_TINT);
  return [Excel.FillPattern.solid, color, color].join("|");
}

// Zero-based indexes of the lines a band covers.
function bandLines(count: number): number[] {
  const lines: number[] = [];
  for (let line = FIRST_BAND; line < count; line += BAND_STEP) lines.push(line);
  return lines;
}

// On means every cell of every banded line already carries exactly this tint;
// one hand-painted cell among them makes the press a fresh band, not a clear.
function alreadyBanded(
  grid: string[][],
  lines: number[],
  axis: PinstripeAxis,
  key: string,
): boolean {
  return lines.every((line) =>
    axis === "rows"
      ? (grid[line] ?? []).every((cell) => cell === key)
      : grid.every((row) => row[line] === key),
  );
}

function lineRange(
  range: Excel.Range,
  axis: PinstripeAxis,
  line: number,
): Excel.Range {
  return axis === "rows" ? range.getRow(line) : range.getColumn(line);
}

// True when the sheet's filter is switched on, has criteria applied, and its
// range shares at least one cell with the selection: a filter elsewhere on
// the sheet, or one with nothing currently filtered, leaves banding alone.
async function filterActive(
  context: Excel.RequestContext,
  range: Excel.Range,
  autoFilter: Excel.AutoFilter,
): Promise<boolean> {
  if (!autoFilter.enabled || !autoFilter.isDataFiltered) return false;
  const filterRange = autoFilter.getRange();
  filterRange.load("address");
  await context.sync();
  return intersects(
    parseAddress(range.address).address,
    parseAddress(filterRange.address).address,
  );
}

// Zero-based row indexes, local to the selection, that are hidden right now -
// by the filter this was asked about, or by hand underneath it: one lineRange
// load per row, since a single row's own rowHidden is never the null Excel
// answers for a mixed multi-row band.
async function hiddenSelectionRows(
  context: Excel.RequestContext,
  range: Excel.Range,
  count: number,
): Promise<Set<number>> {
  const rows = Array.from({ length: count }, (_unused, line) =>
    lineRange(range, "rows", line),
  );
  rows.forEach((row) => row.load("rowHidden"));
  await context.sync();
  const hidden = new Set<number>();
  rows.forEach((row, line) => {
    if (row.rowHidden) hidden.add(line);
  });
  return hidden;
}

// Zero-based indexes of the lines a band covers when a filter hides some of
// the selection: the same every-second rule as bandLines, over the lines the
// filter left showing, so the bands fall where Excel's own banded-table style
// would put them.
function bandVisibleLines(hidden: Set<number>, count: number): number[] {
  const visible: number[] = [];
  for (let line = 0; line < count; line += 1) {
    if (!hidden.has(line)) visible.push(line);
  }
  return bandLines(visible.length).map((index) => visible[index]!);
}

// Zero-based row indexes a filter is hiding within the selection, or null to
// band by position: the axis, the host version and an active, intersecting
// filter all have to agree first. Shares the sync requestFills already
// queued for the fills themselves.
async function filterHiddenRows(
  context: Excel.RequestContext,
  range: Excel.Range,
  axis: PinstripeAxis,
  count: number,
): Promise<Set<number> | null> {
  // A host below ExcelApi 1.9 has no autoFilter to ask, and a column band
  // never needs one: a filter cannot hide a column.
  const autoFilter =
    axis === "rows" && hostSupports("1.9") ? range.worksheet.autoFilter : null;
  autoFilter?.load("enabled,isDataFiltered");
  await context.sync();
  if (!autoFilter || !(await filterActive(context, range, autoFilter))) {
    return null;
  }
  return hiddenSelectionRows(context, range, count);
}

function done(axis: PinstripeAxis, lines: number, banded: boolean): string {
  const noun = axis === "rows" ? "row" : "column";
  const plural = lines === 1 ? noun : `${noun}s`;
  return `${STAGE}: ${String(lines)} ${plural} ${banded ? "banded" : "cleared"}`;
}

/**
 * Bands every second row or column of the selection, or clears the bands when
 * they all still carry exactly the band tint. pls,fix Undo puts the fills that
 * were there before back, either way.
 */
export async function applyPinstripes(axis: PinstripeAxis): Promise<string> {
  // An overlay's fills are not the modeller's, and the overlay hands them back
  // on the way out: banding over one would return our tint as their
  // formatting. The store itself names whichever overlay is holding them.
  requireNoOverlayOwner(STAGE);
  return Excel.run(async (context) => {
    // The cap answers before the fills are asked for: a clicked column header
    // is a million cells, and a grid read of those would freeze the pane.
    const range = await withinCap(
      context,
      await selectedSingleRange(context, STAGE),
      STAGE,
    );
    // The address only matters for a row band checking itself against a
    // filter; a column band never needs it, since a filter never hides one.
    range.load(
      axis === "rows" ? "rowCount,columnCount,address" : "rowCount,columnCount",
    );
    await context.sync();

    const count = axis === "rows" ? range.rowCount : range.columnCount;
    if (count < MIN_LINES) {
      const noun = axis === "rows" ? "rows" : "columns";
      throw new Error(`${STAGE} need at least two ${noun} in the selection.`);
    }
    // A protected sheet refuses every fill. A band is a reading aid, not an
    // edit worth an error dialog, so it is skipped instead and said so - and
    // the undo slot is left holding the previous action.
    if (await sheetProtected(context, range.worksheet)) {
      return protectedNote(STAGE);
    }

    const properties = requestFills(range);
    const hidden = await filterHiddenRows(context, range, axis, count);

    const key = bandKey();
    const lines = hidden ? bandVisibleLines(hidden, count) : bandLines(count);
    const banded = alreadyBanded(fillGrid(properties), lines, axis, key);
    await captureUndo(context, range);

    for (const line of lines) {
      const band = lineRange(range, axis, line);
      if (banded) band.format.fill.clear();
      else applyFillKey(band, key);
    }
    return paintSync(context, STAGE, done(axis, lines.length, !banded));
  });
}
