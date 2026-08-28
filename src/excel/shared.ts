// Public building blocks the Excel adapter's sections share: formatting preset
// and number-format names, the selection summary shape, and address parsing.
// No Office.js state lives here; every export is a type or a pure function.

export type PresetName = "title" | "header" | "input" | "formula" | "result";
export type NumberFormatName = "whole" | "decimal" | "currency" | "percent";

export interface SelectionSummary {
  address: string;
  cells: number;
  formulas: number;
  errors: number;
  blanks: number;
}

// Range addresses arrive sheet-qualified; worksheet.getRange wants the local part.
export function parseAddress(address: string): {
  sheet: string;
  address: string;
} {
  const cut = address.lastIndexOf("!");
  if (cut < 0) return { sheet: "", address };
  return {
    sheet: address.slice(0, cut).replace(/^'|'$/g, "").replace(/''/g, "'"),
    address: address.slice(cut + 1),
  };
}
