// Pinstripes: every second row (or column) of the selection banded in a very
// light brand tint, so a wide comps or schedule grid reads across without a
// ruler. A second press over bands that are all still exactly this tint takes
// them off again.
//
// Owns: the band tint, the on/off decision read off the fills themselves, and
// the paint. Invariant: it never paints over an overlay's fills - the audit
// stripes and the linked-cell highlight each hand back what they covered, and
// a band on top of one would be given back as the modeller's own formatting.

import { auditOverlayOn } from "./audit";
import { applyFillKey, selectedSingleRange, withinCap } from "./internal";
import { fillGrid, requestFills } from "./fill-store";
import { linkHighlightOn } from "./link-highlight";
import { paintSync, protectedNote, sheetProtected } from "./protection";
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

// An overlay's fills are not the modeller's, and the overlay hands them back on
// the way out: banding over one would return our tint as their formatting. The
// refusal is worded exactly as the two overlays word it to each other.
function requireNoOverlay(): void {
  if (auditOverlayOn()) {
    throw new Error(`${STAGE}: turn the audit overlay off first`);
  }
  if (linkHighlightOn()) {
    throw new Error(`${STAGE}: turn the linked-cell highlight off first`);
  }
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
  requireNoOverlay();
  return Excel.run(async (context) => {
    // The cap answers before the fills are asked for: a clicked column header
    // is a million cells, and a grid read of those would freeze the pane.
    const range = await withinCap(
      context,
      await selectedSingleRange(context, STAGE),
      STAGE,
    );
    range.load("rowCount,columnCount");
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
    await context.sync();

    const key = bandKey();
    const lines = bandLines(count);
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
