// The helper block a bar chart is drawn from: the label/low/high triples read
// off the selection, the number format its two value columns wear, and the
// write that lands the block immediately right of the selection.
//
// Owns: everything the tornado and the football field do identically before
// their chart is added. Invariant: nothing is written before the target block
// is known empty and the sheet is known to take writes, so a refusal never
// spends a pls,fix Undo slot.

import { requireEmptyBlock, SHEET_COLUMNS } from "./internal";
import { protectedNote, sheetProtected, syncWrite } from "./protection";
import { captureUndo } from "./undo";
import { type CellValue } from "../model";

/** Label, low and high: the three columns both charts read and write. */
export const CHART_BLOCK_COLUMNS = 3;
const PLAIN_FORMAT = "General";

export function isFiniteNumber(value: CellValue | undefined): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

/** One row of the block's source: what it is called, and its two ends. */
export interface LabelledRange {
  label: string;
  low: number;
  high: number;
}

/** What the reader refuses, worded by the chart that asked. */
export interface TripleRules {
  minRows: number;
  tooFew: string;
  notNumbers: string;
  /** A cap of the reader's own; omitted when the caller counted the rows. */
  rowCap?: { max: number; message: string };
}

/**
 * The selection read as label, low and high. The header row is optional and
 * detected, not declared: a first row whose two number cells hold no numbers is
 * a set of column titles, never a driver or a valuation method.
 */
export function readTriples(
  grid: CellValue[][],
  rules: TripleRules,
): LabelledRange[] {
  const first = grid[0] ?? [];
  const headed = !isFiniteNumber(first[1]) && !isFiniteNumber(first[2]);
  const body = headed ? grid.slice(1) : grid;
  if (body.length < rules.minRows) throw new Error(rules.tooFew);
  if (rules.rowCap && body.length > rules.rowCap.max) {
    throw new Error(rules.rowCap.message);
  }

  return body.map((row) => {
    const [label, low, high] = row;
    if (!isFiniteNumber(low) || !isFiniteNumber(high)) {
      throw new Error(rules.notNumbers);
    }
    return { label: label === null ? "" : String(label), low, high };
  });
}

// The block's numbers share the selection's unit, so both of its value columns
// wear the format of the first data row's low cell; a headed selection has one
// row above that.
export function valueFormat(formats: string[][], rowCount: number): string {
  return formats[formats.length - rowCount]?.[1] ?? PLAIN_FORMAT;
}

// The header row and the label column stay plain; the two value columns carry
// the selection's format, which the chart reads off them.
function blockFormats(format: string, rowCount: number): string[][] {
  return [
    Array.from({ length: CHART_BLOCK_COLUMNS }, () => PLAIN_FORMAT),
    ...Array.from({ length: rowCount }, () => [PLAIN_FORMAT, format, format]),
  ];
}

/** Refuses a selection with no room for the three columns beside it. */
export function requireRoomBeside(range: Excel.Range, stage: string): void {
  const past =
    range.columnIndex + range.columnCount + CHART_BLOCK_COLUMNS > SHEET_COLUMNS;
  if (past) throw new Error(`${stage}: no room to the right of the selection`);
}

/** What a chart's helper block is made of. */
export interface HelperBlock {
  stage: string;
  headers: readonly string[];
  /** One row per bar: its label and the two numbers the chart plots. */
  rows: (string | number)[][];
  /** The number format the two value columns wear. */
  format: string;
}

/**
 * Writes the helper block immediately right of the selection and hands it back
 * for the chart to be built on. A block with anything in it, or a sheet that
 * refuses writes, is refused before pls,fix Undo captures anything.
 */
export async function writeHelperBlock(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  range: Excel.Range,
  block: HelperBlock,
): Promise<Excel.Range> {
  const target = sheet.getRangeByIndexes(
    range.rowIndex,
    range.columnIndex + range.columnCount,
    block.rows.length + 1,
    CHART_BLOCK_COLUMNS,
  );
  await requireEmptyBlock(
    context,
    target,
    `${block.stage}: cells to the right of the selection are not empty`,
  );
  // Asked before the capture, not after: a refusal that had spent an Undo slot
  // would push the modeller's last real action off the five-deep stack.
  if (await sheetProtected(context, sheet)) {
    throw new Error(protectedNote(block.stage));
  }
  await captureUndo(context, target);

  target.values = [[...block.headers], ...block.rows];
  target.numberFormat = blockFormats(block.format, block.rows.length);
  // A locked cell refuses the helper block by name, before a chart is added
  // over a block that never landed.
  await syncWrite(context, block.stage);
  return target;
}
