// The five brand presets as one write over a range's format, plus the plain
// body look a label keeps. Owns what "Title", "Header", "Input", "Formula" and
// "Result" mean in a workbook; the Model formatting buttons and every template
// block paint through this one function, so the two can never drift apart.
// Not barrel-exported: the pane reaches presets through `applyPreset`.

import { type PresetName } from "./shared";
import { activeTheme, getActiveSettings } from "../settings";

/** A preset, or the plain body text a template's label cells wear. */
export type PresetLook = PresetName | "label";

export function applyPresetFormat(
  format: Excel.RangeFormat,
  look: PresetLook,
): void {
  const theme = activeTheme();
  format.font.name = getActiveSettings().font;
  format.font.size = 10;
  format.font.bold = false;
  format.font.italic = false;
  format.font.color = theme.formulaFont;
  format.fill.clear();
  format.horizontalAlignment = Excel.HorizontalAlignment.left;
  format.verticalAlignment = Excel.VerticalAlignment.center;

  switch (look) {
    // The reset above is the whole look: body text on no fill.
    case "label":
    case "formula":
      break;
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
