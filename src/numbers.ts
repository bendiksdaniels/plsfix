// House number style per language: thousands grouped by a space (lv, ru) or a
// comma (en), the decimal always a point, the currency after the amount (lv,
// ru) or before it (en). Excel's own separators are an application setting the
// pane can only read, so format codes stay locale-neutral (#,##0) and the
// language decides currency placement plus every number the pane writes as
// text. Pure.

export type Language = "lv" | "en" | "ru";

export interface LanguageOption {
  code: Language;
  label: string;
}

export const LANGUAGES: readonly LanguageOption[] = [
  { code: "lv", label: "Latviešu" },
  { code: "en", label: "English" },
  { code: "ru", label: "Русский" },
];

export interface NumberStyle {
  /** What separates thousands in text the pane writes. */
  grouping: " " | ",";
  decimal: ".";
  /** `1 094 417 €` rather than `€ 1,094,417`. */
  currencyAfter: boolean;
}

const STYLES: Record<Language, NumberStyle> = {
  lv: { grouping: " ", decimal: ".", currencyAfter: true },
  en: { grouping: ",", decimal: ".", currencyAfter: false },
  ru: { grouping: " ", decimal: ".", currencyAfter: true },
};

export function isLanguage(value: unknown): value is Language {
  return value === "lv" || value === "en" || value === "ru";
}

export function numberStyle(language: Language): NumberStyle {
  return STYLES[language];
}

/**
 * A number as text in the house style, at most `fractionDigits` decimals with
 * trailing zeros dropped: 1 094 417.5 in Latvian, 1,094,417.5 in English.
 */
export function formatAmount(
  value: number,
  language: Language,
  fractionDigits = 1,
): string {
  const style = numberStyle(language);
  const scale = 10 ** fractionDigits;
  const rounded = Math.round(Math.abs(value) * scale) / scale;
  const [whole = "0", fraction = ""] = rounded
    .toFixed(fractionDigits)
    .split(".");
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, style.grouping);
  const decimals = fraction.replace(/0+$/, "");
  const sign = value < 0 && rounded !== 0 ? "-" : "";
  return `${sign}${grouped}${decimals ? style.decimal + decimals : ""}`;
}

/** An Excel format body for money: the symbol before or after the digits. */
export function currencyFormat(
  symbol: string,
  language: Language,
  digits: string,
): string {
  if (!symbol) return digits;
  return numberStyle(language).currencyAfter
    ? `${digits} ${symbol}`
    : `${symbol} ${digits}`;
}

/** What Excel shows for 1094417.5 with the separators it is set to. */
export function separatorSample(decimal: string, thousands: string): string {
  return `1${thousands}094${thousands}417${decimal}5`;
}

// Excel reports the separators the operating system's locale data holds, not
// the ones a keyboard types: the Latvian and Russian locales spell their
// thousands separator as U+00A0 NO-BREAK SPACE, and other locales use the
// narrow (U+202F) or thin (U+2009) one. All of them show as the gap the house
// style asks for, so both sides are canonicalised before they are compared -
// state Excel gives back is never exact-matched.
const SPACES = /[\u00A0\u202F\u2009\u2007]/g;

function canonical(separator: string): string {
  return separator.replace(SPACES, " ");
}

/** Whether Excel's separators already show this language's house style. */
export function separatorsMatch(
  decimal: string,
  thousands: string,
  language: Language,
): boolean {
  const style = numberStyle(language);
  return (
    canonical(decimal) === style.decimal &&
    canonical(thousands) === style.grouping
  );
}

/** The two separators an application shows numbers with, read not written. */
export interface NumberSeparators {
  decimal: string;
  thousands: string;
}

// Office.js always hands back a number's text with the invariant comma and
// point, whatever the application is set to show: optional leading sign or
// opening parenthesis, digits grouped by comma, an optional point fraction, an
// optional trailing percent, then a closing parenthesis or the trailing
// spaces an accounting format pads a positive number with to line up under a
// parenthesised one. Anything else - a currency symbol, a date - is not this
// shape and is left alone.
const PLAIN_NUMBER = /^[-+(]?\d+(?:,\d{3})*(?:\.\d+)?%?[)\s]*$/;

/**
 * Rewrites a number's invariant-separator text to the separators an
 * application actually shows, in one pass so a source already carrying the
 * target characters is never touched twice (tasks/lessons.md, 2026-08-27:
 * never exact-match state Excel gives back - the same rule applies to
 * rewriting it). A text that is not a plain formatted number, or an
 * application already on the invariant separators, comes back unchanged.
 */
export function localizeNumberText(
  text: string,
  separators: NumberSeparators,
): string {
  const { decimal, thousands } = separators;
  if (decimal === "." && thousands === ",") return text;
  if (!PLAIN_NUMBER.test(text)) return text;
  return text.replace(/[,.]/g, (mark) => (mark === "," ? thousands : decimal));
}
