// The action dispatch table: turns a [data-action] id (button click or ribbon
// shortcut) into the excel/pane call it makes. Called through the shared
// guard by main.ts's button loop and its styles-delete confirm. Office.js
// only reaches here through ../excel.

import {
  addCagrLabel,
  applyBorderCycle,
  applyColumnWidthCycle,
  applyDecimalStep,
  applyFillCycle,
  applyFontColorCycle,
  applyNumberCycle,
  applyNumberFormat,
  applyPreset,
  applyRowHeightCycle,
  applyRowStyleCycle,
  applySignFlip,
  autocolorSelection,
  clearFormats,
  fastFillAuto,
  formatSelectedChart,
  insertCagr,
  insertColorKey,
  insertConsistentRounding,
  insertTemplate,
  insertTornado,
  insertWaterfall,
  markCopySource,
  pasteSpecial,
  pastePreserveFormulas,
  scaleSelection,
  toggleIfErrorGuard,
  undoLastAction,
  unpivotSelection,
  type NumberFormatName,
  type PresetName,
  type TraceDirection,
} from "../excel";
import type { NumberCycleFamily, RowStyleKind } from "../cycles";
import { runFind } from "./find-panel";
import { applyPaintSlot, capturePaintSlot } from "./paint-slots";
import { prepareShare } from "./share-panel";
import { deleteStyles, scanStyles } from "./styles-panel";
import { startTrace, toggleAudit } from "./trace-panel";
import { insertTocSheet, scanNames } from "./workbook-tab";

export async function dispatch(action: string): Promise<string> {
  if (action.startsWith("style-")) {
    await applyPreset(action.replace("style-", "") as PresetName);
  } else if (action.startsWith("paint-capture-")) {
    return capturePaintSlot(Number(action.replace("paint-capture-", "")));
  } else if (action.startsWith("paint-apply-")) {
    return applyPaintSlot(Number(action.replace("paint-apply-", "")));
  } else if (action.startsWith("number-")) {
    await applyNumberFormat(action.replace("number-", "") as NumberFormatName);
  } else if (action.startsWith("cycle-number-")) {
    await applyNumberCycle(
      action.replace("cycle-number-", "") as NumberCycleFamily,
    );
  } else if (action.startsWith("cycle-row-")) {
    await applyRowStyleCycle(action.replace("cycle-row-", "") as RowStyleKind);
  } else if (action.startsWith("template-")) {
    return insertTemplate(action.replace("template-", ""));
  } else {
    switch (action) {
      case "cycle-fill":
        await applyFillCycle();
        break;
      case "cycle-font":
        await applyFontColorCycle();
        break;
      case "cycle-border":
        await applyBorderCycle();
        break;
      case "cycle-row-height":
        await applyRowHeightCycle();
        break;
      case "cycle-col-width":
        await applyColumnWidthCycle();
        break;
      case "clear-formats":
        await clearFormats();
        break;
      case "fill-right":
        await fastFillAuto("right");
        break;
      case "fill-down":
        await fastFillAuto("down");
        break;
      case "if-error":
        await toggleIfErrorGuard();
        break;
      case "autocolor":
        return autocolorSelection();
      case "insert-color-key":
        return insertColorKey();
      case "divide-1000":
        await scaleSelection(0.001);
        break;
      case "multiply-1000":
        await scaleSelection(1000);
        break;
      case "cagr":
        await insertCagr();
        break;
      case "sign-flip":
        await applySignFlip();
        break;
      case "dec-more":
        await applyDecimalStep(1);
        break;
      case "dec-less":
        await applyDecimalStep(-1);
        break;
      case "undo":
        return `Restored ${await undoLastAction()}`;
      case "copy-source":
        return `Copy source: ${await markCopySource()}`;
      case "paste-values":
        await pasteSpecial("values");
        break;
      case "paste-formats":
        await pasteSpecial("formats");
        break;
      case "paste-transpose":
        await pasteSpecial("transpose");
        break;
      case "paste-exact":
        await pastePreserveFormulas();
        break;
      case "chart-waterfall":
        return insertWaterfall();
      case "tornado":
        return insertTornado();
      case "unpivot":
        return unpivotSelection();
      case "write-rounded":
        return insertConsistentRounding();
      case "chart-format":
        await formatSelectedChart();
        return "Chart restyled to your brand";
      case "chart-cagr":
        return addCagrLabel();
      case "audit-toggle":
        return toggleAudit();
      case "trace-precedents":
      case "trace-dependents":
        return startTrace(
          action.replace("trace-", "") as TraceDirection,
          false,
        );
      case "insert-toc":
        return insertTocSheet();
      case "find":
        return runFind();
      case "scan-names":
        return scanNames();
      case "styles-scan":
        return scanStyles();
      case "styles-delete":
        return deleteStyles();
      case "share-prepare":
        return prepareShare();
      default:
        throw new Error(`Unknown action: ${action}`);
    }
  }

  return "Selection updated";
}
