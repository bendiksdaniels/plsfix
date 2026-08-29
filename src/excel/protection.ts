// Whether a sheet will take a write at all. Excel refuses every write to a
// locked cell of a protected sheet with a bare AccessDenied, which reaches the
// pane as an error dialog over something the modeller only asked to be painted.
//
// Owns: the protection read (Excel.WorksheetProtection, ExcelApi 1.2) and the
// line a flow reports instead. Invariant: a flow that only paints never throws
// on protection - it says which sheet refused it and changes nothing.

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
