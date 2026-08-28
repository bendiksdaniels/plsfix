// Reshaping the selection. Unpivot writes its long table to a new worksheet, so
// no cell in the model is overwritten and SMT Undo has nothing to capture.

import { selectedSingleRange, withinCap } from "./internal";
import { type CellValue } from "../model";
import { unpivot } from "../reshape";
import { activeTheme, getActiveSettings } from "../settings";

const UNPIVOT_SHEET = "Unpivot";
const UNPIVOT_HEADERS = ["Row", "Column", "Value"];
const UNPIVOT_SHEET_LIMIT = 99;
const UNPIVOT_FONT_SIZE = 10;

// Sheet names are case-insensitive in Excel, so a workbook that already holds
// an "unpivot" gets "Unpivot 2" rather than an itemAlreadyExists on add().
function freeSheetName(taken: string[], base: string): string {
  const used = new Set(taken.map((name) => name.toLowerCase()));
  if (!used.has(base.toLowerCase())) return base;

  for (let suffix = 2; suffix <= UNPIVOT_SHEET_LIMIT; suffix += 1) {
    const candidate = `${base} ${suffix}`;
    if (!used.has(candidate.toLowerCase())) return candidate;
  }
  throw new Error("unpivot: this workbook already holds 99 Unpivot sheets");
}

// Wide to long, the shape a pivot table or a database wants: the header row and
// the first column become the two key columns, one line per populated cell.
export async function unpivotSelection(): Promise<string> {
  return Excel.run(async (context) => {
    const range = await withinCap(
      context,
      await selectedSingleRange(context, "unpivot"),
      "Unpivot",
    );
    range.load("values");
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const rows = unpivot(range.values as CellValue[][]);
    if (rows.length === 0) {
      throw new Error("unpivot: the selection holds no values");
    }

    const name = freeSheetName(
      sheets.items.map((item) => item.name),
      UNPIVOT_SHEET,
    );
    const sheet = sheets.add(name);
    const width = UNPIVOT_HEADERS.length;
    const header = sheet.getRangeByIndexes(0, 0, 1, width);
    header.values = [UNPIVOT_HEADERS];
    sheet.getRangeByIndexes(1, 0, rows.length, width).values = rows as (
      string | number | boolean
    )[][];

    const theme = activeTheme();
    const table = sheet.getRangeByIndexes(0, 0, rows.length + 1, width).format;
    table.font.name = getActiveSettings().font;
    table.font.size = UNPIVOT_FONT_SIZE;
    header.format.font.bold = true;
    header.format.font.color = theme.titleText;
    header.format.fill.color = theme.titleFill;

    sheet.activate();
    await context.sync();

    return `Unpivot: ${rows.length} rows on ${name}`;
  });
}
