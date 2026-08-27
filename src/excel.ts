import {
  analyzeGrid,
  type CellValue,
  makeFormatGrid,
  scaleCells,
  wrapFormulasWithIfError,
} from "./model";

export type PresetName = "title" | "header" | "input" | "formula" | "result";
export type NumberFormatName = "whole" | "decimal" | "currency" | "percent";

export interface SelectionSummary {
  address: string;
  cells: number;
  formulas: number;
  errors: number;
  blanks: number;
}

const numberFormats: Record<NumberFormatName, string> = {
  whole: "#,##0;[Red](#,##0);-",
  decimal: "#,##0.0;[Red](#,##0.0);-",
  currency: "€ #,##0;[Red](€ #,##0);-",
  percent: "0.0%;[Red](0.0%);-",
};

export async function inspectSelection(): Promise<SelectionSummary> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address,rowCount,columnCount,formulas,values");
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
    const { format } = range;

    format.font.name = "Aptos";
    format.font.size = 10;
    format.font.bold = false;
    format.font.italic = false;
    format.font.color = "#172A3A";
    format.fill.clear();
    format.horizontalAlignment = Excel.HorizontalAlignment.left;
    format.verticalAlignment = Excel.VerticalAlignment.center;

    switch (name) {
      case "title":
        format.fill.color = "#172A3A";
        format.font.color = "#FFFFFF";
        format.font.size = 15;
        format.font.bold = true;
        format.rowHeight = 25;
        break;
      case "header": {
        format.fill.color = "#E8EEF3";
        format.font.bold = true;
        const bottom = format.borders.getItem(Excel.BorderIndex.edgeBottom);
        bottom.style = Excel.BorderLineStyle.continuous;
        bottom.color = "#8293A1";
        bottom.weight = Excel.BorderWeight.thin;
        break;
      }
      case "input":
        format.font.color = "#0057B8";
        break;
      case "formula":
        format.font.color = "#172A3A";
        break;
      case "result": {
        format.fill.color = "#E8F3EC";
        format.font.bold = true;
        const top = format.borders.getItem(Excel.BorderIndex.edgeTop);
        top.style = Excel.BorderLineStyle.double;
        top.color = "#2F6B4F";
        break;
      }
    }

    await context.sync();
  });
}

export async function clearFormats(): Promise<void> {
  await Excel.run(async (context) => {
    context.workbook
      .getSelectedRange()
      .clear(Excel.ClearApplyTo.formats);
    await context.sync();
  });
}

export async function applyNumberFormat(name: NumberFormatName): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount");
    await context.sync();
    range.numberFormat = makeFormatGrid(
      range.rowCount,
      range.columnCount,
      numberFormats[name],
    );
    await context.sync();
  });
}

export async function fastFill(direction: "right" | "down"): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount,formulas");
    await context.sync();

    if (direction === "right" && range.rowCount !== 1) {
      throw new Error("Fill right needs a single-row selection.");
    }
    if (direction === "down" && range.columnCount !== 1) {
      throw new Error("Fill down needs a single-column selection.");
    }

    const source = range.getCell(0, 0);
    source.load("formulas");
    await context.sync();
    const sourceFormula = source.formulas[0]?.[0] as CellValue | undefined;
    if (typeof sourceFormula !== "string" || !sourceFormula.startsWith("=")) {
      throw new Error("The first selected cell must contain a formula.");
    }

    range.copyFrom(source, Excel.RangeCopyType.formulas);
    await context.sync();
  });
}

export async function addIfError(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("formulas");
    await context.sync();
    range.formulas = wrapFormulasWithIfError(
      range.formulas as CellValue[][],
    ) as (string | number | boolean)[][];
    await context.sync();
  });
}

export async function scaleSelection(factor: 1000 | 0.001): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("formulas");
    await context.sync();
    range.formulas = scaleCells(
      range.formulas as CellValue[][],
      factor,
    ) as (string | number | boolean)[][];
    await context.sync();
  });
}

export async function autocolorSelection(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount,formulas,values");
    await context.sync();

    if (range.rowCount * range.columnCount > 5_000) {
      throw new Error("Autocolor supports up to 5,000 selected cells at once.");
    }

    const formulas = range.formulas as CellValue[][];
    const values = range.values as CellValue[][];

    for (let row = 0; row < range.rowCount; row += 1) {
      for (let column = 0; column < range.columnCount; column += 1) {
        const formula = formulas[row]?.[column] ?? null;
        const value = values[row]?.[column] ?? null;
        if (value === null || value === "") continue;

        const cell = range.getCell(row, column);
        if (typeof formula === "string" && formula.startsWith("=")) {
          cell.format.font.color = formula.includes("[") ? "#17823B" : "#172A3A";
        } else {
          cell.format.font.color = "#0057B8";
        }
      }
    }

    await context.sync();
  });
}
