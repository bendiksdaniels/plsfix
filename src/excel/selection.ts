// Selection formatting: the brand presets, the house number formats and the
// format eraser. Every mutating action captures pls,fix Undo first and writes
// into every area of the selection, so a ctrl-clicked pair of blocks is treated
// as one job. The format cycles live next door in format-cycles.ts.

import { cappedAreas, selectedAreas } from "./areas";
import { numberFormat, SELECTION_CELL_CAP } from "./internal";
import {
  type NumberFormatName,
  type PresetName,
  type SelectionSummary,
} from "./shared";
import { captureUndoAreas } from "./undo";
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

// The look itself, so one area and five areas go through the same code.
function paintPreset(range: Excel.Range, name: PresetName): void {
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
}

export async function applyPreset(name: PresetName): Promise<void> {
  await Excel.run(async (context) => {
    const areas = await selectedAreas(context, "Formatting");
    await captureUndoAreas(context, areas);
    for (const area of areas) paintPreset(area, name);
    await context.sync();
  });
}

export async function clearFormats(): Promise<void> {
  await Excel.run(async (context) => {
    const areas = await selectedAreas(context, "Clearing formats");
    await captureUndoAreas(context, areas);
    for (const area of areas) area.clear(Excel.ClearApplyTo.formats);
    await context.sync();
  });
}

export async function applyNumberFormat(name: NumberFormatName): Promise<void> {
  await Excel.run(async (context) => {
    const areas = await cappedAreas(context, "Number formatting");
    for (const area of areas) area.load("rowCount,columnCount");
    await context.sync();
    await captureUndoAreas(context, areas);

    const format = numberFormat(name);
    for (const area of areas) {
      area.numberFormat = makeFormatGrid(
        area.rowCount,
        area.columnCount,
        format,
      );
    }
    await context.sync();
  });
}
