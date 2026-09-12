// "Precedents of selection": the Smart Track precedents step run over every
// formula cell of one selected block, grouped by source cell and unioned.
// Owns the batching (one sync for the whole selection) and the union select.
// Invariant: read-only - it selects the answer, it never writes a cell.

import { selectedSingleRange, syncTolerating } from "./internal";
import { parseAddress } from "./shared";
import { queueTrace, traceAreas, type TraceArea } from "./trace";
import { cellAddress } from "../find";
import { type CellValue, isFormula } from "../model";

// One batch asks Excel about every cell at once, and a batch is one request:
// fifty formula cells is already a long answer to render as chips.
export const MULTI_TRACE_CELL_CAP = 50;
// Where halving a failed batch stops paying: below this, splitting costs a
// sync per level for the same answers one cell at a time would give.
const FLAT_RETRY_CELLS = 4;

const STAGE = "Precedents of selection";
const OVER_CAP = `${STAGE} handles up to ${String(MULTI_TRACE_CELL_CAP)} cells at once.`;

/** What one cell of the selection reads from. */
export interface CellPrecedents {
  /** The source cell's own address, without its sheet. */
  cell: string;
  areas: TraceArea[];
}

export interface MultiTraceResult {
  /** The selection the walk started from, sheet-qualified. */
  origin: string;
  groups: CellPrecedents[];
  /** Every area any source cell reads from, each one once. */
  areas: TraceArea[];
  formulaCells: number;
  /** Cells of the selection that hold no formula, so were never asked about. */
  skipped: number;
}

interface SourceCell {
  address: string;
  range: Excel.Range;
}

/** Each cell its own batch, in order: what a small failed batch is worth. */
async function resolveOneByOne(
  context: Excel.RequestContext,
  cells: SourceCell[],
): Promise<CellPrecedents[]> {
  const groups: CellPrecedents[] = [];
  for (const cell of cells)
    groups.push(...(await resolveGroups(context, [cell])));
  return groups;
}

/**
 * One batch for the whole list. Excel answers a cell with no precedents by
 * throwing ItemNotFound, which fails the entire batch and names no cell, so a
 * failed batch is halved and each half asked again - and once a failing batch
 * is down to FLAT_RETRY_CELLS, straight through cell by cell: the cells that
 * read from nothing are the ones still failing alone. A model row where every
 * cell reads from something costs exactly one round trip.
 */
async function resolveGroups(
  context: Excel.RequestContext,
  cells: SourceCell[],
): Promise<CellPrecedents[]> {
  const queued = cells.map((cell) => ({
    cell,
    found: queueTrace(cell.range, "precedents"),
  }));
  if (await syncTolerating(context, Excel.ErrorCodes.itemNotFound)) {
    return queued.map((entry) => ({
      cell: entry.cell.address,
      areas: traceAreas(entry.found),
    }));
  }

  const first = cells[0];
  if (cells.length === 1 && first) return [{ cell: first.address, areas: [] }];
  if (cells.length <= FLAT_RETRY_CELLS) return resolveOneByOne(context, cells);

  const half = Math.ceil(cells.length / 2);
  const head = await resolveGroups(context, cells.slice(0, half));
  const tail = await resolveGroups(context, cells.slice(half));
  return [...head, ...tail];
}

/** Every area any source cell named, first mention wins. */
function unionAreas(groups: CellPrecedents[]): TraceArea[] {
  const seen = new Map<string, TraceArea>();
  for (const group of groups) {
    for (const area of group.areas) {
      const key = `${area.sheet}!${area.address}`;
      if (!seen.has(key)) seen.set(key, area);
    }
  }
  return Array.from(seen.values());
}

function formulaCellsOf(
  range: Excel.Range,
  grid: CellValue[][],
  sheetRow: number,
  sheetColumn: number,
): { cells: SourceCell[]; total: number } {
  const cells: SourceCell[] = [];
  let total = 0;
  grid.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      total += 1;
      if (!isFormula(value)) return;
      cells.push({
        address: cellAddress(sheetRow + rowIndex, sheetColumn + columnIndex),
        range: range.getCell(rowIndex, columnIndex),
      });
    });
  });
  return { cells, total };
}

/**
 * The direct precedents of every formula cell in the selection. Five syncs on
 * a selection that all reads from something: the area count, the cell count
 * before the grid, the grid, the one batch of trace calls, and the select.
 */
export async function tracePrecedentsOfSelection(): Promise<MultiTraceResult> {
  return Excel.run(async (context) => {
    const range = await selectedSingleRange(context, STAGE);
    // Counted before the grid is loaded: a clicked column header is a million
    // cells, and the batch would ask Excel about every one of them.
    range.load("address,cellCount,rowIndex,columnIndex");
    await context.sync();

    if (range.cellCount < 0 || range.cellCount > MULTI_TRACE_CELL_CAP) {
      throw new Error(OVER_CAP);
    }
    const origin = range.address;
    const sheetName = parseAddress(origin).sheet;
    range.load("formulasR1C1");
    await context.sync();

    const { cells, total } = formulaCellsOf(
      range,
      range.formulasR1C1 as CellValue[][],
      range.rowIndex,
      range.columnIndex,
    );
    if (cells.length === 0) {
      return { origin, groups: [], areas: [], formulaCells: 0, skipped: total };
    }

    const groups = await resolveGroups(context, cells);
    const areas = unionAreas(groups);
    // A precedent on another sheet is listed, never selected: selecting it
    // would take the modeller off the block they are reviewing.
    const here = areas
      .filter((area) => area.sheet === "" || area.sheet === sheetName)
      .map((area) => area.address);
    if (here.length > 0) {
      range.worksheet.getRanges(here.join(",")).select();
      await context.sync();
    }

    return {
      origin,
      groups,
      areas,
      formulaCells: cells.length,
      skipped: total - cells.length,
    };
  });
}
