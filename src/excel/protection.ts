// What the host refuses when a flow writes, and what the pane says instead of
// the bare string office.js hands back: a locked cell of a protected sheet
// (AccessDenied), and a write covering only part of a merged cell
// (InvalidOperation).
//
// Owns: the protection read (Excel.WorksheetProtection, ExcelApi 1.2) and both
// translations. Invariant: a flow that only paints never throws on protection -
// it says which sheet refused it and changes nothing; a flow that edits does
// throw, but with its own stage and a way out in the message.

import { hostSupports } from "./internal";

const PROTECTION_API_SET = "1.2";

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

/** The line the pane shows instead of an error when a paint was refused. */
export function protectedNote(stage: string): string {
  return `${stage}: this sheet is protected, nothing was changed`;
}

/**
 * Runs a batch of edits. A refusal comes back named and staged: which sheet is
 * protected, or that the selection cuts a merged cell. Everything else travels
 * untouched.
 */
export async function syncWrite(
  context: Excel.RequestContext,
  stage: string,
): Promise<void> {
  try {
    await context.sync();
  } catch (error) {
    const { code } = error as { code?: string };
    if (code === Excel.ErrorCodes.accessDenied) {
      throw new Error(protectedNote(stage));
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
 */
export async function paintSync(
  context: Excel.RequestContext,
  stage: string,
  done: string,
): Promise<string> {
  try {
    await context.sync();
    return done;
  } catch (error) {
    if ((error as { code?: string }).code === Excel.ErrorCodes.accessDenied) {
      return protectedNote(stage);
    }
    throw error;
  }
}
