// The Excel pane's help copy: one line per section on what it is for, then one
// or two plain sentences per button. Owns the wording only - the section keys
// and the button keys are the pane's own markup (aria-labelledby, data-action
// or id), and the coverage rule lives in copy.ts. Same claims as the Latvian
// manual: a sentence here and its chapter there must not disagree.

import type { HelpCopy } from "./copy";

const tools: HelpCopy = {
  "selection-heading": {
    about: "What the selection holds right now, before you act on it.",
    buttons: {
      "refresh-selection":
        "Reads the selection again, for when the counts look stale after an edit in Excel.",
    },
  },
  "format-heading": {
    about:
      "One look for the whole model: five cell styles, four number formats and three format slots.",
    buttons: {
      "style-title": "Applies the title style from your Brand palette.",
      "style-header":
        "Applies the header style, for the label row above a block of numbers.",
      "style-input":
        "Applies the input style, for numbers that are typed in by hand.",
      "style-formula": "Applies the formula style, for calculated cells.",
      "style-result":
        "Applies the result style, for the totals a block adds up to.",
      "clear-formats":
        "Strips the formatting off the selection and leaves the values alone.",
      "number-whole": "Whole numbers, no decimals.",
      "number-decimal": "Numbers with one decimal place.",
      "number-currency":
        "Currency format. The symbol and where it sits follow the Brand tab.",
      "number-percent": "Percentages with one decimal place.",
      "paint-capture-1":
        "Remembers the active cell's number format, font, fill, alignment and borders in slot 1.",
      "paint-capture-2": "The same, into slot 2.",
      "paint-capture-3": "The same, into slot 3.",
      "paint-apply-1": "Paints what slot 1 remembers over the selection.",
      "paint-apply-2": "Paints what slot 2 remembers over the selection.",
      "paint-apply-3": "Paints what slot 3 remembers over the selection.",
    },
  },
  "cycles-heading": {
    about:
      "Buttons that step. Press the same one again and the format moves to the next variant.",
    buttons: {
      "cycle-number-general":
        "Steps the number format: 1,234, then 1,234.0, then 1,234.00.",
      "cycle-number-date":
        "Steps the date format: 31.12.2026, then Dec-26, then 2026.",
      "cycle-number-currency":
        "Steps the currency format: whole, then one decimal, then thousands.",
      "cycle-number-percent":
        "Steps the percentage format: 12.3%, then 12%, then 12.34%.",
      "cycle-number-multiple": "Steps the multiple format: 1.5x, then 1.50x.",
      "cycle-row-title":
        "Steps a title row through its looks. Row styles run per row, up to 500 at a time.",
      "cycle-row-result":
        "Steps a result row through its looks, for the totals a block adds up to.",
      "cycle-row-item":
        "Steps an item row through its looks, for the lines inside a block.",
      "cycle-fill": "Steps the fill color through your Brand palette.",
      "cycle-font": "Steps the font color through your Brand palette.",
      "cycle-border":
        "Steps the borders: bottom rule, total, result, box, then grid.",
      "cycle-row-height": "Steps the row height: 15, 18, 21, 24, then 30 pt.",
      "cycle-col-width": "Steps the column width: 64, 80, 96, 120, then 48 pt.",
    },
  },
  "tools-heading": {
    about:
      "The everyday edits on a selection: fill, paste, guard, color and quick maths.",
    buttons: {
      undo: "Puts back what the last pls,fix action overwrote. Excel's own Ctrl+Z never sees add-in writes.",
      "fill-right":
        "Copies the first cell's formula right. One cell alone is sized by the rows beside it.",
      "fill-down":
        "Copies the first cell's formula down. One cell alone is sized by the columns beside it.",
      "if-error":
        "Wraps the selected formulas in IFERROR(..., 0). Press again to strip the guard off.",
      autocolor:
        "Colors the selection by content: typed numbers, formulas, cross-sheet links, external links and hardcodes.",
      "insert-color-key":
        "Drops the color legend on the sheet, starting at the active cell.",
      "write-rounded":
        "Writes =PLSFIX.ROUND beside the selected row or column, so the rounded parts add up to the rounded total.",
      unpivot:
        "Rewrites a cross-tab as Row, Column and Value lines on a new sheet. The source is left alone.",
      "copy-source":
        "Marks the selection as the source the four paste buttons read from.",
      "paste-values": "Pastes the marked source as values only.",
      "paste-formats": "Pastes the marked source's formatting only.",
      "paste-exact":
        "Pastes the marked source's formulas exactly, references unchanged.",
      "paste-transpose": "Pastes the marked source with its rows as columns.",
      "divide-1000":
        "Divides the selection by a thousand. Works on typed numbers and on formulas.",
      "multiply-1000":
        "Multiplies the selection by a thousand. Works on typed numbers and on formulas.",
      cagr: "Writes the compound annual growth rate just past the selected row or column.",
      "sign-flip": "Flips the sign of the selection.",
      "dec-more": "Adds one decimal place to the selection.",
      "dec-less": "Takes one decimal place off the selection.",
    },
  },
  "audit-heading": {
    about:
      "Two ways to check formulas: an overlay, and a walk along the chain.",
    buttons: {
      "audit-toggle":
        "Stripes cells whose formula matches their neighbours and reddens the ones that break the pattern.",
      "trace-precedents":
        "Lists the cells the active cell reads from. Click an address to jump there and keep tracing.",
      "trace-dependents":
        "Lists the cells that read from the active cell, the same way round.",
      "trace-back": "Steps the trace panel back to the cell you came from.",
    },
  },
  "charts-heading": {
    about:
      "Charts built from a selection, and a restyle for a chart you already have.",
    buttons: {
      "chart-waterfall":
        "Builds a bridge chart from a label and value table. First and last rows are the opening and closing totals.",
      tornado:
        "Builds a sensitivity chart from a driver, low and high table, ranked by swing.",
      "chart-format":
        "Restyles the selected chart to your palette: fonts, series colors and legend, gridlines off.",
      "chart-cagr":
        "Writes the selected series' growth rate at the top right of the selection.",
    },
  },
};

const brand: HelpCopy = {
  "preview-heading": {
    about:
      "A live sample of how the settings below look in a model. Nothing here is a button.",
    buttons: {},
  },
  "palette-heading": {
    about: "The seven colors every pls,fix style and autocolor draws from.",
    buttons: {
      "reset-brand": "Puts all seven colors back to the pls,fix defaults.",
    },
  },
  "logo-heading": {
    about:
      "Reads the colors out of a logo. The image is read in the pane and never uploaded.",
    buttons: {},
  },
  "settings-heading": {
    about:
      "Font, language, currency and on-edit autocolor, plus handing the whole lot to a colleague.",
    buttons: {
      "export-brand":
        "Copies every Brand setting to the clipboard, ready for someone else's Import JSON.",
    },
  },
};

const workbook: HelpCopy = {
  "sheets-heading": {
    about:
      "Every sheet in the file: click a name to go there, the circle hides or shows it.",
    buttons: {
      "refresh-sheets":
        "Reads the sheet list again, after sheets were added, renamed or deleted.",
    },
  },
  "find-heading": {
    about: "One search across the whole workbook, hidden sheets included.",
    buttons: {
      find: "Searches values, formula text, defined names, sheet names and comments on every sheet.",
    },
  },
  "styles-heading": {
    about:
      "Custom cell styles pile up in a file and slow it down. This finds the ones nothing wears.",
    buttons: {
      "styles-scan":
        "Lists the custom cell styles no cell in this workbook uses.",
      "styles-delete":
        "Deletes the styles the scan found. A sheet too large to scan blocks the delete.",
    },
  },
  "toc-heading": {
    about: "A contents sheet that links to every visible sheet.",
    buttons: {
      "insert-toc":
        "Builds the contents sheet, or rewrites it from scratch when it is already there.",
    },
  },
  "share-heading": {
    about: "One pass over the file before it leaves your hands.",
    buttons: {
      "share-prepare":
        "Puts every visible sheet back at A1, then reports what a reader would still find. Nothing is deleted.",
    },
  },
  "names-heading": {
    about:
      "Defined names left pointing at deleted cells cause errors later. This clears them out.",
    buttons: {
      "scan-names":
        "Lists the defined names that point at #REF!, that is, at deleted cells.",
      "delete-names": "Deletes the broken names the scan found.",
    },
  },
};

const links: HelpCopy = {
  "export-heading": {
    about:
      "Send a range or a chart to PowerPoint as an object you can refresh later.",
    buttons: {
      "export-selection":
        "Sends the selected range as a linked picture. The selection must be one unbroken block.",
      "export-chart":
        "Sends the selected chart as a linked picture. With none selected, the sheet's only chart is taken.",
      "export-table":
        "Sends the selected range as a table that stays editable in PowerPoint.",
    },
  },
  "links-heading": {
    about: "Every link this workbook feeds, and what you can do to them.",
    buttons: {
      "push-selected":
        "Renders the ticked links again and sends them to the relay.",
      "push-all": "Renders every link in this workbook and sends them all.",
      "go-to-source":
        "Jumps to the range or chart the ticked link was exported from.",
      "remove-link":
        "Drops the link from this workbook. A picture already in a deck stays where it is.",
    },
  },
  "linkkey-heading": {
    about: "The shared key that pairs this workbook's exports with a deck.",
    buttons: {
      "generate-key":
        "Makes a new key. Decks still holding the old one stop seeing new exports.",
      "copy-key":
        "Copies the whole key, ready to paste into PowerPoint > Settings.",
      "reveal-key":
        "Shows the key in full. The panel otherwise shows its two ends only.",
      "forget-key":
        "Deletes the key from this computer. Exports already pushed stay on the relay.",
    },
  },
};

export const EXCEL_HELP: HelpCopy = {
  ...tools,
  ...brand,
  ...workbook,
  ...links,
};
