// Selection formatting: the brand presets, the house number formats and the
// format eraser. Every mutating action captures pls,fix Undo first and writes
// into every area of the selection, so a ctrl-clicked pair of blocks is treated
// as one job. The format cycles live next door in format-cycles.ts.

import { cappedAreas, selectedAreas } from "./areas";
import { numberFormat, overCap } from "./internal";
import { syncWrite } from "./protection";
import {
  type NumberFormatName,
  type PresetName,
  type SelectionSummary,
} from "./shared";
import { captureUndoAreas } from "./undo";
import { analyzeGrid, type CellValue, makeFormatGrid } from "../model";
import { activeTheme, getActiveSettings } from "../settings";

// Every area's metrics added together: a ctrl-clicked selection is one card in
// the pane, not two, and a passive click must never be an error.
function totals(areas: Excel.Range[]): Omit<SelectionSummary, "address"> {
  return areas.reduce(
    (sum, area) => {
      const part = analyzeGrid(
        area.formulas as CellValue[][],
        area.values as CellValue[][],
      );
      return {
        cells: sum.cells + part.cells,
        formulas: sum.formulas + part.formulas,
        errors: sum.errors + part.errors,
        blanks: sum.blanks + part.blanks,
      };
    },
    { cells: 0, formulas: 0, errors: 0, blanks: 0 },
  );
}

export async function inspectSelection(): Promise<SelectionSummary> {
  return Excel.run(async (context) => {
    const areas = await selectedAreas(context, "Selection");
    for (const area of areas) area.load("address,cellCount");
    await context.sync();

    const address = areas.map((area) => area.address).join(", ");
    const counts = areas.map((area) => area.cellCount);
    const cells = counts.some((count) => count < 0)
      ? -1
      : counts.reduce((total, count) => total + count, 0);
    // Over the cap the pane shows the address and count only (metrics as "—")
    // instead of asking the host for two full-column grids on a passive click.
    if (overCap(cells)) {
      return { address, cells, formulas: -1, errors: -1, blanks: -1 };
    }

    for (const area of areas) area.load("formulas,values");
    await context.sync();

    return { address, ...totals(areas) };
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
    await syncWrite(context, "Formatting");
  });
}

export async function clearFormats(): Promise<void> {
  await Excel.run(async (context) => {
    const areas = await selectedAreas(context, "Clearing formats");
    await captureUndoAreas(context, areas);
    for (const area of areas) area.clear(Excel.ClearApplyTo.formats);
    await syncWrite(context, "Clearing formats");
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
    await syncWrite(context, "Number formatting");
  });
}
