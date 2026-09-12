// Ribbon commands and keyboard shortcuts (shared runtime): one command table
// mapping every PLSFIX_* id to the action it runs, then a promise chain per
// command instead of the shared guard, since these can fire with the pane
// closed. Office.js only reaches here through ../excel.

import {
  addCagrLabel,
  applyBorderCycle,
  applyColumnWidthCycle,
  applyDecimalStep,
  applyFillCycle,
  applyFontColorCycle,
  applyNumberCycle,
  applyRowHeightCycle,
  applyRowStyleCycle,
  applySignFlip,
  autocolorSelection,
  fastFillAuto,
  formatSelectedChart,
  insertCagr,
  insertConsistentRounding,
  insertTornado,
  insertWaterfall,
  lastUndoSkipped,
  markCopySource,
  pasteSpecial,
  pastePreserveFormulas,
  scaleSelection,
  toggleIfErrorGuard,
  undoLastAction,
  unpivotSelection,
} from "../excel";
import { describeError } from "../ui/report";
import { focusFind } from "./find-panel";
import { capturePaintSlot, applyPaintSlot } from "./paint-slots";
import { prepareShare } from "./share-panel";
import {
  APP_VERSION,
  refreshSelection,
  renderActionState,
  toast,
} from "./shared";
import { focusStyles } from "./styles-panel";
import { startTrace, toggleAudit } from "./trace-panel";
import { insertTocSheet } from "./workbook-tab";

interface CommandEvent {
  completed: () => void;
}

// The two ids whose whole answer is in the pane: Super Find lists what it
// found in the Workbook tab and the style scrubber lists the unused styles
// there, so a keystroke with the pane shut would draw where nobody is
// looking. Everything else on the shortcut card writes to the workbook -
// cells, formats, a chart, a sheet, or the selection the trace pair jumps -
// and is read there with the pane open or not.
const NEEDS_PANE = new Set(["PLSFIX_FIND", "PLSFIX_STYLES_SCAN"]);

// Office chrome, not a workbook action: a host without the API, or one that
// refuses, still lets the action behind it run.
async function showPane(): Promise<void> {
  await Promise.resolve(Office.addin?.showAsTaskpane()).catch(() => undefined);
}

export function registerCommands(): void {
  if (!Office.actions?.associate) return;

  const commands: Record<string, () => Promise<unknown>> = {
    PLSFIX_AUTOCOLOR: autocolorSelection,
    PLSFIX_AUDIT: toggleAudit,
    // With the pane closed there is nothing to read, so the keystroke jumps to
    // the first result instead.
    PLSFIX_TRACE_PRE: () => startTrace("precedents", true),
    PLSFIX_TRACE_DEP: () => startTrace("dependents", true),
    PLSFIX_FILLRIGHT: () => fastFillAuto("right"),
    PLSFIX_FILLDOWN: () => fastFillAuto("down"),
    PLSFIX_IFERROR: toggleIfErrorGuard,
    PLSFIX_SCALEUP: () => scaleSelection(1000),
    PLSFIX_SCALEDOWN: () => scaleSelection(0.001),
    PLSFIX_UNDO: undoLastAction,
    PLSFIX_COPYSRC: markCopySource,
    PLSFIX_PASTE_VALUES: () => pasteSpecial("values"),
    PLSFIX_PASTE_FORMATS: () => pasteSpecial("formats"),
    PLSFIX_PASTE_EXACT: pastePreserveFormulas,
    PLSFIX_PASTE_TRANSPOSE: () => pasteSpecial("transpose"),
    PLSFIX_CAGR: insertCagr,
    PLSFIX_ROUND: insertConsistentRounding,
    PLSFIX_SIGN: applySignFlip,
    PLSFIX_DEC_MORE: () => applyDecimalStep(1),
    PLSFIX_DEC_LESS: () => applyDecimalStep(-1),
    PLSFIX_CYC_GENERAL: () => applyNumberCycle("general"),
    PLSFIX_CYC_DATE: () => applyNumberCycle("date"),
    PLSFIX_CYC_CURRENCY: () => applyNumberCycle("currency"),
    PLSFIX_CYC_PERCENT: () => applyNumberCycle("percent"),
    PLSFIX_CYC_MULTIPLE: () => applyNumberCycle("multiple"),
    PLSFIX_CYC_TITLE: () => applyRowStyleCycle("title"),
    PLSFIX_CYC_RESULT: () => applyRowStyleCycle("result"),
    PLSFIX_CYC_ITEM: () => applyRowStyleCycle("item"),
    PLSFIX_CYC_FILL: applyFillCycle,
    PLSFIX_CYC_FONT: applyFontColorCycle,
    PLSFIX_CYC_BORDER: applyBorderCycle,
    PLSFIX_CYC_ROWH: applyRowHeightCycle,
    PLSFIX_CYC_COLW: applyColumnWidthCycle,
    PLSFIX_PAINT_CAP1: () => capturePaintSlot(1),
    PLSFIX_PAINT_CAP2: () => capturePaintSlot(2),
    PLSFIX_PAINT_CAP3: () => capturePaintSlot(3),
    PLSFIX_PAINT_APP1: () => applyPaintSlot(1),
    PLSFIX_PAINT_APP2: () => applyPaintSlot(2),
    PLSFIX_PAINT_APP3: () => applyPaintSlot(3),
    PLSFIX_WATERFALL: insertWaterfall,
    PLSFIX_TORNADO: insertTornado,
    PLSFIX_CHARTFMT: formatSelectedChart,
    PLSFIX_CHART_CAGR: addCagrLabel,
    PLSFIX_UNPIVOT: unpivotSelection,
    PLSFIX_TOC: insertTocSheet,
    PLSFIX_SHARE: prepareShare,
    PLSFIX_FIND: focusFind,
    PLSFIX_STYLES_SCAN: focusStyles,
  };

  for (const [id, run] of Object.entries(commands)) {
    const start = NEEDS_PANE.has(id) ? showPane : () => Promise.resolve();
    Office.actions.associate(id, (event?: CommandEvent) => {
      void start()
        // Wrapped, never `.then(run)`: that hands the action the resolved
        // value as an argument, and half this table takes arguments.
        .then(() => run())
        .then(() => refreshSelection())
        .catch((error: unknown) => {
          const { message, details } = describeError(
            error,
            { host: "Excel", version: APP_VERSION },
            id,
          );
          toast.show(message, "error", details);
        })
        .finally(() => {
          // Drain the skip flag so a later pane toast cannot inherit it.
          lastUndoSkipped();
          renderActionState();
          event?.completed();
        });
    });
  }

  Office.actions.associate("PLSFIX_SHOWPANE", (event?: CommandEvent) => {
    void showPane().finally(() => event?.completed());
  });
}
