// Human names for what a format cycle just landed on, in the exact step
// order src/cycles.ts's own tables build - never Excel's own format code or
// a hex colour, which reads as line noise to the modeller a toast is for.
// Split out of cycles.ts (already at its own 400-line cap) rather than
// pushed further over it. Pure: no Office.js, no DOM.

import { currencyFormat } from "./numbers";
import type { BrandSettings } from "./settings";
import type { NumberCycleFamily, RowStyleKind } from "./cycles";

// buildNumberCycles' own step order for the three families a currency never
// touches. A currency symbol places before or after the digits by language,
// so its own labels are built fresh from the live settings below instead.
const STATIC_NUMBER_LABELS: Record<
  Exclude<NumberCycleFamily, "currency">,
  string[]
> = {
  general: ["1,234", "1,234.0", "1,234.00"],
  percent: ["12.3%", "12%", "12.34%"],
  multiple: ["0.0x", "0.00x"],
  date: ["dd.mm.yyyy", "mmm-yy", "yyyy"],
};

// The trailing comma in the cycle's third currency format divides the
// displayed value by a thousand - Excel's own scaling trick for a model kept
// in whole currency but shown in thousands.
function currencyLabels(settings: BrandSettings): string[] {
  const { currency, language } = settings;
  return [
    currencyFormat(currency, language, "1,234"),
    currencyFormat(currency, language, "1,234.0"),
    `${currencyFormat(currency, language, "1,234")} (thousands)`,
  ];
}

/** Same shape and step order as buildNumberCycles(settings). */
export function numberCycleLabels(
  settings: BrandSettings,
): Record<NumberCycleFamily, string[]> {
  return { ...STATIC_NUMBER_LABELS, currency: currencyLabels(settings) };
}

/** buildRowStyleCycles' own step order. */
export const ROW_STYLE_LABELS: Record<RowStyleKind, string[]> = {
  title: ["Filled", "Header rule", "Accent rule"],
  result: ["Filled", "Double rule", "Accent block"],
  item: ["Plain", "Header tint", "Result tint"],
};

// buildFillCycle's own step order. headerFill/resultFill are theme tints of
// primary/accent (deriveTheme), not palette slots of their own, so they get
// a descriptive name rather than the settings key nobody edits directly.
export const FILL_CYCLE_LABELS = [
  "Header tint",
  "Result tint",
  "Accent",
  "Primary",
  "none",
];

// buildFontCycle's own step order, named the way the Brand tab's own palette
// rows already label these four slots (taskpane.html).
export const FONT_CYCLE_LABELS = [
  "Formulas",
  "Inputs",
  "Cross-sheet links",
  "Accent",
  "Primary",
];

// buildBorderCycle's own step order: nothing, a rule under the row, a
// heavier one, the double rule a total takes, a box, then the full grid.
export const BORDER_CYCLE_LABELS = [
  "none",
  "underline",
  "heavy underline",
  "double underline",
  "box",
  "grid",
];

const ALIGN_LABELS: Record<string, string> = {
  Left: "left",
  Center: "centre",
  Right: "right",
  General: "general",
};

const UNDERLINE_LABELS: Record<string, string> = {
  Single: "single",
  Double: "double",
  None: "none",
};

export function alignLabel(value: string): string {
  return ALIGN_LABELS[value] ?? value.toLowerCase();
}

export function underlineLabel(value: string): string {
  return UNDERLINE_LABELS[value] ?? value.toLowerCase();
}
