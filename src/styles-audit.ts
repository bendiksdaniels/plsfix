// The style scrubber's rule: which of a workbook's cell styles no cell wears.
// A model that has been copied between workbooks for years carries hundreds of
// them, and a bloated style table is what makes such a file slow to open.
// Pure: src/excel/styles.ts reads the style table and the cells themselves.

export interface WorkbookStyle {
  name: string;
  builtIn: boolean;
}

// Excel's own styles (Normal, Comma, Percent, ...) cannot be deleted and would
// come back anyway, so only the styles someone added are ever offered. Excel
// keeps style names unique regardless of case, so a cell reporting a differently
// cased spelling is still that style being used.
export function unusedStyles(
  all: WorkbookStyle[],
  used: Set<string>,
): string[] {
  const worn = new Set(Array.from(used, (name) => name.toLowerCase()));
  return all
    .filter((style) => !style.builtIn && !worn.has(style.name.toLowerCase()))
    .map((style) => style.name)
    .sort((a, b) => a.localeCompare(b));
}
