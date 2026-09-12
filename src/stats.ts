// The comps statistics block: which row of a selected comps table is a header,
// which of its columns hold numbers, and the six live formulas each of those
// columns gets under the table. Pure - a grid in, A1 formula strings out.
// Invariant: every formula spans exactly the data rows, never the header row,
// and the block's first column is a label column, never a statistic.

import { cellAddress } from "./find";
import { type CellValue } from "./model";

/** The six rows a comps page reads, in that order. */
export const STATS_LABELS = [
  "Min",
  "25th percentile",
  "Median",
  "Mean",
  "75th percentile",
  "Max",
] as const;

export const STATS_ROWS = STATS_LABELS.length;
/** The blank row the block keeps between itself and the table above it. */
export const STATS_GAP_ROWS = 1;

const MIN_DATA_ROWS = 2;
// The label column: it carries the six row names, so it never carries a
// statistic even when a modeller has numbers in it.
const LABEL_COLUMN = 0;

/** What a preset the block paints with may be; null leaves the cell alone. */
export type StatsLook = "label" | "formula" | null;

export interface CompsBlock {
  /** Zero-based sheet row of the first and last data row. */
  firstDataRow: number;
  lastDataRow: number;
  /** Zero-based sheet column of the block's first (label) column. */
  firstColumn: number;
  /** Whether the selection's first row was read as column titles. */
  headed: boolean;
  /** Column offsets from firstColumn that hold at least one number. */
  numericColumns: number[];
}

function isNumber(value: CellValue | undefined): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function isText(value: CellValue | undefined): boolean {
  return typeof value === "string" && value.trim() !== "";
}

/**
 * A first row of column titles, not a company. A comps table usually leaves the
 * label column's own title blank, so the rule is "no numbers, some text" rather
 * than "every cell is text" - which the tornado's own detector already follows.
 */
export function isHeaderRow(row: CellValue[]): boolean {
  return !row.some(isNumber) && row.some(isText);
}

function numericOffsets(dataRows: CellValue[][]): number[] {
  const width = dataRows.reduce((most, row) => Math.max(most, row.length), 0);
  const offsets: number[] = [];
  for (let column = LABEL_COLUMN + 1; column < width; column += 1) {
    if (dataRows.some((row) => isNumber(row[column]))) offsets.push(column);
  }
  return offsets;
}

/**
 * The selection read as a comps table. `rowIndex` and `columnIndex` are the
 * selection's own zero-based corner, so the spans the formulas carry are the
 * addresses Excel will resolve them against.
 */
export function readCompsBlock(
  grid: CellValue[][],
  rowIndex: number,
  columnIndex: number,
): CompsBlock {
  const first = grid[0] ?? [];
  const headed = isHeaderRow(first);
  const dataRows = headed ? grid.slice(1) : grid;
  if (dataRows.length < MIN_DATA_ROWS) {
    throw new Error("Comps stats need a table with at least two data rows.");
  }

  const numericColumns = numericOffsets(dataRows);
  if (numericColumns.length === 0) {
    throw new Error("Comps stats need at least one column of numbers.");
  }

  const firstDataRow = rowIndex + (headed ? 1 : 0);
  return {
    firstDataRow,
    lastDataRow: firstDataRow + dataRows.length - 1,
    firstColumn: columnIndex,
    headed,
    numericColumns,
  };
}

// "C6:C8": the data rows of one column, as Excel names them.
function span(block: CompsBlock, offset: number): string {
  const column = block.firstColumn + offset;
  const from = cellAddress(block.firstDataRow, column);
  const to = cellAddress(block.lastDataRow, column);
  return `${from}:${to}`;
}

function columnFormulas(range: string): string[] {
  return [
    `=MIN(${range})`,
    `=PERCENTILE.INC(${range},0.25)`,
    `=MEDIAN(${range})`,
    `=AVERAGE(${range})`,
    `=PERCENTILE.INC(${range},0.75)`,
    `=MAX(${range})`,
  ];
}

/**
 * The six rows the block writes: the labels down its first column, one live
 * formula per numeric column, and "" everywhere else - `Range.formulas` is
 * assigned as one whole grid, so every position needs an entry.
 */
export function statsGrid(block: CompsBlock, columnCount: number): string[][] {
  const formulas = new Map(
    block.numericColumns.map((offset) => [
      offset,
      columnFormulas(span(block, offset)),
    ]),
  );
  return STATS_LABELS.map((label, row) =>
    Array.from({ length: columnCount }, (_unused, column) => {
      if (column === LABEL_COLUMN) return label as string;
      return formulas.get(column)?.[row] ?? "";
    }),
  );
}

/**
 * The look each cell of the block wears: the plain label look down the first
 * column (the Item row style's own first variant), the formula look on the
 * statistics, and nothing at all on the columns that stay empty.
 */
export function statsPresets(
  block: CompsBlock,
  columnCount: number,
): StatsLook[][] {
  const numeric = new Set(block.numericColumns);
  const row: StatsLook[] = Array.from(
    { length: columnCount },
    (_unused, column) => {
      if (column === LABEL_COLUMN) return "label";
      return numeric.has(column) ? "formula" : null;
    },
  );
  return STATS_LABELS.map(() => [...row]);
}
