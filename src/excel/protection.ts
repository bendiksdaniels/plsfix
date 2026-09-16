// What the host refuses when a flow writes, and what the pane says instead of
// the bare string office.js hands back: a locked cell of a protected sheet
// (AccessDenied), a protected workbook structure refusing a change to the
// sheet list (AccessDenied too), and a write covering only part of a merged
// cell (InvalidOperation).
//
// Owns: the protection reads (Excel.WorksheetProtection, ExcelApi 1.2;
// Excel.WorkbookProtection, ExcelApi 1.7) and the translations. Invariant: a
// flow that only paints never throws on protection - it says which sheet
// refused it and changes nothing; a flow that edits does throw, but with its
// own stage and a way out in the message.

import { hostSupports } from "./internal";
import { commitUndo, discardUndo } from "./undo";

const PROTECTION_API_SET = "1.2";
const WORKBOOK_PROTECTION_API_SET = "1.7";

/**
 * True when the sheet refuses writes. A host too old to be asked answers false:
 * the write is then attempted, and paintSync catches the refusal.
 */
export async function sheetProtected(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
): Promise<boolean> {
  const protection = (sheet as unknown as Record<string, unknown>).protection;
  if (!protection || !hostSupports(PROTECTION_API_SET)) return false;
  sheet.protection.load("protected");
  await context.sync();
  return sheet.protection.protected;
}

/**
 * The line the pane shows instead of an error when a paint was refused. A
 * batch of more than one area (a ctrl-clicked selection) may have already
 * landed some of them before the host refused another, so it is worded as
 * partly protected rather than promising an untouched sheet.
 */
export function protectedNote(stage: string, areas = 1): string {
  return areas > 1
    ? `${stage}: this sheet is partly protected, the areas before the locked one were changed, the rest were not`
    : `${stage}: this sheet is protected, nothing was changed`;
}

/**
 * The same for the sheet list: hiding, showing or moving a sheet is refused
 * while the workbook's structure is protected, and Excel says only
 * "AccessDenied" - which reads as a locked cell if it is not told apart.
 */
export function structureNote(stage: string): string {
  return `${stage}: this workbook's structure is protected, nothing was changed`;
}

/**
 * True when the sheet list is locked. A host too old to be asked answers
 * false: the write is then attempted, and syncWrite catches the refusal.
 */
export async function structureProtected(
  context: Excel.RequestContext,
): Promise<boolean> {
  const workbook = context.workbook as unknown as Record<string, unknown>;
  if (!workbook.protection || !hostSupports(WORKBOOK_PROTECTION_API_SET)) {
    return false;
  }
  context.workbook.protection.load("protected");
  await context.sync();
  return context.workbook.protection.protected;
}

/**
 * Runs a batch of edits. A refusal comes back named and staged: which sheet is
 * protected, or that the selection cuts a merged cell. Everything else travels
 * untouched. A flow writing the sheet list rather than cells passes
 * `structureNote` as `refused`, so AccessDenied is worded for what it wrote.
 * `areas` is the number of ranges the batch wrote in one go (a ctrl-clicked
 * selection): callers writing more than one pass their count on, so a locked
 * area among unlocked ones is worded as partly protected, not untouched.
 * Settles the pending pls,fix Undo entry either way: committed once the write
 * lands, discarded before every throw so a refused write never spends a slot.
 */
export async function syncWrite(
  context: Excel.RequestContext,
  stage: string,
  refused: (stage: string, areas?: number) => string = protectedNote,
  areas = 1,
): Promise<void> {
  try {
    await context.sync();
    commitUndo();
  } catch (error) {
    discardUndo();
    const { code } = error as { code?: string };
    if (code === Excel.ErrorCodes.accessDenied) {
      throw new Error(refused(stage, areas));
    }
    if (code === Excel.ErrorCodes.invalidOperation) {
      throw new Error(
        `${stage}: Excel refused this write. Select whole merged cells, not part of one.`,
      );
    }
    throw error;
  }
}

/**
 * Runs a paint batch. A protected sheet or a locked cell comes back as the
 * note rather than as a rejection; every other failure travels untouched.
 * Settles the pending pls,fix Undo entry either way, the same as syncWrite.
 */
export async function paintSync(
  context: Excel.RequestContext,
  stage: string,
  done: string,
): Promise<string> {
  try {
    await context.sync();
    commitUndo();
    return done;
  } catch (error) {
    discardUndo();
    if ((error as { code?: string }).code === Excel.ErrorCodes.accessDenied) {
      return protectedNote(stage);
    }
    throw error;
  }
}
