// Selection formatting: presets, number formats and format-cycling (fill, font
// color, row style, number). Every mutating action captures SMT Undo first and
// enforces the selection cell cap before touching the grid.

import {
  numberFormat,
  SELECTION_CELL_CAP,
  selectionWithinCap,
} from "./internal";
import {
  type NumberFormatName,
  type PresetName,
  type SelectionSummary,
} from "./shared";
import { captureUndo } from "./undo";
import {
  type BorderSpec,
  buildFillCycle,
  buildFontCycle,
  buildNumberCycles,
  buildRowStyleCycles,
  type CellStyle,
  CLEAR_FILL,
  matchStyleIndex,
  nextInCycle,
  type NumberCycleFamily,
  type RowStyleKind,
  type StyleSpec,
} from "../cycles";
import { analyzeGrid, type CellValue, makeFormatGrid } from "../model";
import { activeTheme, getActiveSettings } from "../settings";

export async function inspectSelection(): Promise<SelectionSummary> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address,cellCount");
    await context.sync();

    // Over the cap the pane shows the address and count only (metrics as "—")
    // instead of asking the host for two full-column grids on a passive click.
    if (range.cellCount > SELECTION_CELL_CAP) {
      return {
        address: range.address,
        cells: range.cellCount,
        formulas: -1,
        errors: -1,
        blanks: -1,
      };
    }

    range.load("formulas,values");
    await context.sync();

    const summary = analyzeGrid(
      range.formulas as CellValue[][],
      range.values as CellValue[][],
    );

    return {
      address: range.address,
      ...summary,
    };
  });
}

export async function applyPreset(name: PresetName): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    await captureUndo(context, range);
    const { format } = range;

    const theme = activeTheme();
    format.font.name = getActiveSettings().font;
    format.font.size = 10;
    format.font.bold = false;
    format.font.italic = false;
    format.font.color = theme.formulaFont;
    format.fill.clear();
    format.horizontalAlignment = Excel.HorizontalAlignment.left;
    format.verticalAlignment = Excel.VerticalAlignment.center;

    switch (name) {
      case "title":
        format.fill.color = theme.titleFill;
        format.font.color = theme.titleText;
        format.font.size = 15;
        format.font.bold = true;
        format.rowHeight = 25;
        break;
      case "header": {
        format.fill.color = theme.headerFill;
        format.font.bold = true;
        const bottom = format.borders.getItem(Excel.BorderIndex.edgeBottom);
        bottom.style = Excel.BorderLineStyle.continuous;
        bottom.color = theme.headerBorder;
        bottom.weight = Excel.BorderWeight.thin;
        break;
      }
      case "input":
        format.font.color = theme.inputFont;
        break;
      case "formula":
        format.font.color = theme.formulaFont;
        break;
      case "result": {
        format.fill.color = theme.resultFill;
        format.font.bold = true;
        const top = format.borders.getItem(Excel.BorderIndex.edgeTop);
        top.style = Excel.BorderLineStyle.double;
        top.color = theme.resultBorder;
        break;
      }
    }

    await context.sync();
  });
}

export async function clearFormats(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    await captureUndo(context, range);
    range.clear(Excel.ClearApplyTo.formats);
    await context.sync();
  });
}

export async function applyNumberFormat(name: NumberFormatName): Promise<void> {
  await Excel.run(async (context) => {
    const range = await selectionWithinCap(context, "Number formatting");
    range.load("rowCount,columnCount");
    await context.sync();
    await captureUndo(context, range);

    range.numberFormat = makeFormatGrid(
      range.rowCount,
      range.columnCount,
      numberFormat(name),
    );
    await context.sync();
  });
}

// Cycle state lives in the cell: every run reads the active cell and steps once.
function readFill(cell: Excel.Range): string {
  const { fill } = cell.format;
  // Excel reports white for unfilled cells, so the pattern decides.
  return fill.pattern === Excel.FillPattern.none
    ? CLEAR_FILL
    : fill.color.toUpperCase();
}

function readCellStyle(cell: Excel.Range): CellStyle {
  const fill = readFill(cell);
  return {
    fill: fill === CLEAR_FILL ? null : fill,
    fontColor: cell.format.font.color.toUpperCase(),
    bold: cell.format.font.bold,
  };
}

function applyBorder(
  format: Excel.RangeFormat,
  index: Excel.BorderIndex,
  spec: BorderSpec | null | undefined,
): void {
  if (spec === undefined) return;
  const border = format.borders.getItem(index);
  if (spec === null) {
    border.style = Excel.BorderLineStyle.none;
    return;
  }
  if (spec.style === "double") {
    border.style = Excel.BorderLineStyle.double;
  } else {
    border.style = Excel.BorderLineStyle.continuous;
    border.weight = Excel.BorderWeight.thin;
  }
  border.color = spec.color;
}

function applyStyleSpec(format: Excel.RangeFormat, spec: StyleSpec): void {
  if (spec.fill === CLEAR_FILL) format.fill.clear();
  else if (spec.fill !== undefined) format.fill.color = spec.fill;
  if (spec.fontColor !== undefined) format.font.color = spec.fontColor;
  if (spec.bold !== undefined) format.font.bold = spec.bold;
  applyBorder(format, Excel.BorderIndex.edgeTop, spec.topBorder);
  applyBorder(format, Excel.BorderIndex.edgeBottom, spec.bottomBorder);
}

export async function applyNumberCycle(
  family: NumberCycleFamily,
): Promise<void> {
  await Excel.run(async (context) => {
    const range = await selectionWithinCap(context, "Format cycling");
    const active = range.getCell(0, 0);
    range.load("rowCount,columnCount");
    active.load("numberFormat");
    await context.sync();
    await captureUndo(context, range);

    const current = active.numberFormat[0]?.[0];
    const next = nextInCycle(
      typeof current === "string" ? current : "",
      buildNumberCycles(getActiveSettings())[family],
    );

    range.numberFormat = makeFormatGrid(
      range.rowCount,
      range.columnCount,
      next,
    );
    await context.sync();
  });
}

export async function applyRowStyleCycle(kind: RowStyleKind): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const active = range.getCell(0, 0);
    range.load("rowCount");
    active.load(
      "format/fill/color,format/fill/pattern,format/font/color,format/font/bold",
    );
    await context.sync();

    const variants = buildRowStyleCycles(getActiveSettings())[kind];
    const index = matchStyleIndex(readCellStyle(active), variants);
    const next = variants[(index + 1) % variants.length];
    await captureUndo(context, range);

    if (next) {
      // Edge borders target the whole range, which would leave interior rows
      // bare in a multi-row selection; row styles are per-row by definition.
      // Refuse absurd heights instead of silently degrading to edge borders.
      if (range.rowCount > 500) {
        throw new Error("Row styles support up to 500 rows at once.");
      }
      if (range.rowCount > 1) {
        for (let row = 0; row < range.rowCount; row += 1) {
          applyStyleSpec(range.getRow(row).format, next);
        }
      } else {
        applyStyleSpec(range.format, next);
      }
    }

    await context.sync();
  });
}

export async function applyFillCycle(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const active = range.getCell(0, 0);
    active.load("format/fill/color,format/fill/pattern");
    await context.sync();

    const next = nextInCycle(
      readFill(active),
      buildFillCycle(getActiveSettings()),
    );
    await captureUndo(context, range);

    if (next === CLEAR_FILL) range.format.fill.clear();
    else range.format.fill.color = next;

    await context.sync();
  });
}

export async function applyFontColorCycle(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const active = range.getCell(0, 0);
    active.load("format/font/color");
    await context.sync();

    const next = nextInCycle(
      active.format.font.color.toUpperCase(),
      buildFontCycle(getActiveSettings()),
    );
    await captureUndo(context, range);

    range.format.font.color = next;
    await context.sync();
  });
}
