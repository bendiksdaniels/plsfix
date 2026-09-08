// Copy/paste actions: marking a copy source, paste special (values/formats/
// transpose) and formula-preserving paste. The source is remembered by sheet id
// and address, not a live Excel reference, so it survives a sheet rename.

import { selectedAreas } from "./areas";
import { syncWrite } from "./protection";
import { parseAddress } from "./shared";
import { captureUndoAreas } from "./undo";

export type PasteMode = "values" | "formats" | "transpose";

const PASTE = "Paste";

interface CopySource {
  sheetId: string;
  address: string;
  label: string;
}

let copySource: CopySource | null = null;

export function copySourceLabel(): string | null {
  return copySource?.label ?? null;
}

export async function markCopySource(): Promise<string> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;
    sheet.load("id");
    range.load("address");
    await context.sync();

    copySource = {
      sheetId: sheet.id,
      address: parseAddress(range.address).address,
      label: range.address,
    };
    return range.address;
  });
}

// Resolved by sheet id, so renaming the source sheet between copy and paste is fine.
async function openCopySource(
  context: Excel.RequestContext,
): Promise<Excel.Range> {
  const source = copySource;
  if (!source) throw new Error("Mark a copy source first.");

  const sheet = context.workbook.worksheets.getItemOrNullObject(source.sheetId);
  sheet.load("isNullObject");
  await context.sync();

  if (sheet.isNullObject) {
    copySource = null;
    throw new Error("The copy source sheet is gone. Mark a new source.");
  }
  return sheet.getRange(source.address);
}

// Read inside the call, never at module scope: office.js defines the enums, and
// the pane also loads in a plain browser where they do not exist yet.
function pasteCopyType(mode: PasteMode): Excel.RangeCopyType {
  switch (mode) {
    case "values":
      return Excel.RangeCopyType.values;
    case "formats":
      return Excel.RangeCopyType.formats;
    case "transpose":
      return Excel.RangeCopyType.all;
  }
}

export async function pasteSpecial(mode: PasteMode): Promise<void> {
  await Excel.run(async (context) => {
    const from = await openCopySource(context);
    const targets = await selectedAreas(context, PASTE);
    from.load("rowCount,columnCount");
    for (const target of targets) target.load("rowCount,columnCount");
    await context.sync();

    // Excel grows a smaller destination to the source shape, so undo has to
    // cover the whole footprint, not just what the user selected - once per
    // area, because every area of the selection takes a copy.
    const transposed = mode === "transpose";
    const rows = transposed ? from.columnCount : from.rowCount;
    const columns = transposed ? from.rowCount : from.columnCount;
    await captureUndoAreas(
      context,
      targets.map((target) =>
        target
          .getCell(0, 0)
          .getResizedRange(
            Math.max(rows, target.rowCount) - 1,
            Math.max(columns, target.columnCount) - 1,
          ),
      ),
    );

    for (const target of targets) {
      target.copyFrom(from, pasteCopyType(mode), false, transposed);
    }
    await syncWrite(context, PASTE);
  });
}

// Excel's own paste rewrites relative references; this one keeps the formula
// text byte for byte, which is what a modeller means by "same formula here".
export async function pastePreserveFormulas(): Promise<void> {
  await Excel.run(async (context) => {
    const from = await openCopySource(context);
    const targets = await selectedAreas(context, PASTE);
    from.load("rowCount,columnCount,formulas");
    await context.sync();

    const destinations = targets.map((target) =>
      target
        .getCell(0, 0)
        .getResizedRange(from.rowCount - 1, from.columnCount - 1),
    );
    await captureUndoAreas(context, destinations);

    for (const destination of destinations)
      destination.formulas = from.formulas;
    await syncWrite(context, PASTE);
  });
}
