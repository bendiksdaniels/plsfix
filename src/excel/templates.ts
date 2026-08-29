// The Templates flow: one ready calculation block written with its top-left
// corner at the active cell. Owns the write - values and formulas in one batch,
// number formats and the brand presets in a second - the refusal when the
// target block is not empty, and the Undo capture that makes it reversible.
// The grids themselves are pure data in `src/template-blocks.ts`.

import { numberFormat, requireEmptyBlock, writeRuns } from "./internal";
import { applyPresetFormat, type PresetLook } from "./presets";
import { captureUndo } from "./undo";
import { buildNumberCycles } from "../cycles";
import { getActiveSettings } from "../settings";
import {
  type CellRef,
  resolveFormula,
  type Template,
  type TemplateCell,
  type TemplateFormat,
  templateById,
} from "../templates";

const GENERAL = "General";

// "multiple" is a cycle format rather than one of the number-format buttons, so
// it comes from the first step of that cycle; the rest are the pane's own.
function cellNumberFormat(format: TemplateFormat | undefined): string {
  if (format === undefined) return GENERAL;
  if (format === "multiple") {
    return buildNumberCycles(getActiveSettings()).multiple[0] ?? GENERAL;
  }
  return numberFormat(format);
}

// `formulas` is assigned as one whole grid, so every position needs an entry;
// "" is what a spacer cell is.
function cellContent(cell: TemplateCell, origin: CellRef): string | number {
  if (cell.f !== undefined) return resolveFormula(cell.f, origin);
  return cell.v ?? "";
}

// One preset write per run of same-kind cells, which is how a template row
// reads: a header row is one run, a period row two or three.
function styleBlock(block: Excel.Range, template: Template): void {
  const looks: PresetLook[][] = template.cells.map((row) =>
    row.map((cell) => cell.kind),
  );
  writeRuns(block, looks, (range, look) =>
    applyPresetFormat(range.format, look),
  );
}

export async function insertTemplate(id: string): Promise<string> {
  const template = templateById(id);
  if (!template) throw new Error(`Unknown template: ${id}`);

  return Excel.run(async (context) => {
    const anchor = context.workbook.getActiveCell();
    anchor.load("rowIndex,columnIndex");
    await context.sync();

    const origin: CellRef = {
      row: anchor.rowIndex,
      column: anchor.columnIndex,
    };
    const block = anchor.worksheet.getRangeByIndexes(
      origin.row,
      origin.column,
      template.rows,
      template.cols,
    );
    await requireEmptyBlock(
      context,
      block,
      `Templates need an empty block of ${template.rows} rows by ${template.cols} columns at the selection.`,
    );
    await captureUndo(context, block);

    block.formulas = template.cells.map((row) =>
      row.map((cell) => cellContent(cell, origin)),
    );
    await context.sync();

    block.numberFormat = template.cells.map((row) =>
      row.map((cell) => cellNumberFormat(cell.format)),
    );
    styleBlock(block, template);
    block.select();
    await context.sync();

    return `Template written: ${template.name} (${template.rows}x${template.cols})`;
  });
}
