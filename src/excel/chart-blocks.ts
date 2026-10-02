// The helper block a bar chart is drawn from: the label/low/high triples read
// off the selection, the number format its two value columns wear, and the
// write that lands the block immediately right of the selection.
//
// Owns: everything the tornado and the football field do identically before
// their chart is added, plus serialised(), the shared queue both flows run
// their whole insert through so a second press meets the first press's
// written block rather than racing it while it is still empty. Invariant:
// nothing is written before the target block is known empty and the sheet is
// known to take writes, so a refusal never spends a pls,fix Undo slot.

import { requireEmptyBlock, SHEET_COLUMNS } from "./internal";
import { protectedNote, sheetProtected, syncWrite } from "./protection";
import { captureUndo, undoLastAction, undoTarget } from "./undo";
import { type CellValue } from "../model";

/** Label, low and high: the three columns both charts read and write. */
export const CHART_BLOCK_COLUMNS = 3;
const PLAIN_FORMAT = "General";

// Resolved, never rejected: one call's failure must not take the next one
// with it. Same shape as exclusive() in link-lock.ts, minus the stage name -
// insertFootballField and insertTornado have nothing else to report.
let queue: Promise<void> = Promise.resolve();

/** Runs `work` once every call queued before it has settled, whatever the
 * outcome of those. Never call it from inside `work`: the nested call would
 * wait for a queue only its own caller can advance. */
export function serialised<T>(work: () => Promise<T>): Promise<T> {
  const task = queue.then(() => work());
  queue = task.then(
    () => undefined,
    () => undefined,
  );
  return task;
}

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

// writeHelperBlock's own commit failing: nothing is a completed pls,fix
// action yet (captureUndo's entry is still pending, or the throw happened
// before it), so there is no Undo entry to unwind - only the cells this
// write's own property setters already put there, on a host (or a test's
// injected refusal) that applies a batch's writes before reporting whether
// the batch that carried them landed. Best effort: the caller is already
// unwinding with the original error, which is what the modeller needs to
// see, never a second one from this cleanup.
async function clearUnwrittenBlock(
  context: Excel.RequestContext,
  target: Excel.Range,
): Promise<void> {
  try {
    target.clear(Excel.ClearApplyTo.contents);
    await context.sync();
  } catch {
    // Best effort only.
  }
}

/**
 * Best-effort rollback for the tornado or the football field once their
 * helper block has already landed for real (writeHelperBlock returned, so
 * pls,fix Undo committed it) and a LATER step - the chart, its placement,
 * its styling - refuses: deletes the chart if one was added, then unwinds
 * the block.
 *
 * The pane has no busy latch over an Excel insert: a ribbon command or a
 * keyboard shortcut can capture another pls,fix action while this one sits
 * between its block write and its later chart syncs, pushing its own entry
 * on top of the stack. Undoing blind would then revert that OTHER action
 * instead of this one, and still leave this block on the sheet - so the
 * block is only undone through the modeller's own Undo route when the
 * stack's newest entry is still its own (`undoTarget() === block.address`,
 * the same address captureUndo's own load put on this proxy); otherwise the
 * stack is left alone and the block is cleared directly, the way
 * clearUnwrittenBlock clears one that was never a completed action at all.
 * Swallows its own failure throughout: the caller is already unwinding with
 * the original error, which is what the modeller needs to see.
 */
export async function cleanupChartInsert(
  context: Excel.RequestContext,
  block: Excel.Range,
  chart: Excel.Chart | undefined,
): Promise<void> {
  try {
    if (chart) {
      chart.delete();
      await context.sync();
    }
  } catch {
    // Best effort only.
  }

  if (undoTarget() === block.address) {
    // A failed restore leaves this entry retryable (undoLastAction's own
    // rule) rather than popped, so falling through to clear the cells here
    // would strand the stack pointing at data that is no longer there.
    try {
      await undoLastAction();
    } catch {
      // Best effort only.
    }
    return;
  }
  await clearUnwrittenBlock(context, block);
}

/**
 * Writes the helper block immediately right of the selection and hands it back
 * for the chart to be built on. A block with anything in it, or a sheet that
 * refuses writes, is refused before pls,fix Undo captures anything. Once the
 * write itself is under way, a refusal at its own commit is rolled back
 * (clearUnwrittenBlock) rather than left as cells with real data and no chart.
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

  try {
    target.values = [[...block.headers], ...block.rows];
    target.numberFormat = blockFormats(block.format, block.rows.length);
    // A locked cell refuses the helper block by name, before a chart is added
    // over a block that never landed.
    await syncWrite(context, block.stage);
  } catch (error) {
    await clearUnwrittenBlock(context, target);
    throw error;
  }
  return target;
}
