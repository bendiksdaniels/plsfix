import "./styles.css";
import type { NumberCycleFamily, RowStyleKind } from "./cycles";
import {
  addCagrLabel,
  activateSheet,
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
  applySlot,
  autocolorSelection,
  captureSlot,
  clearFormats,
  copySourceLabel,
  deleteBrokenNames,
  deleteUnusedStyles,
  fastFillAuto,
  findInWorkbook,
  formatSelectedChart,
  insertCagr,
  insertColorKey,
  insertConsistentRounding,
  insertTornado,
  insertWaterfall,
  insertToc,
  inspectSelection,
  jumpToHit,
  lastUndoSkipped,
  restorePersistedOverlay,
  listBrokenNames,
  listSheets,
  listUnusedStyles,
  markCopySource,
  parseAddress,
  pasteSpecial,
  pastePreserveFormulas,
  readWorkbookBrand,
  scaleSelection,
  selectArea,
  setAutocolorOnEdit,
  setSheetVisibility,
  toggleAuditOverlay,
  toggleIfErrorGuard,
  traceActiveCell,
  undoLastAction,
  undoTarget,
  unpivotSelection,
  writeWorkbookBrand,
  type FindHit,
  type FindResult,
  type NumberFormatName,
  type PresetName,
  type SheetEntry,
  type StyleScan,
  type TraceArea,
  type TraceDirection,
  type TraceResult,
} from "./excel";
import { FIND_HIT_CAP } from "./find";
import { RelayClient, relayBaseUrl } from "./link/relay";
import { officeKeyStore } from "./link/workspace";
import { installLinksTab } from "./pane/links-tab";
import {
  emptySlots,
  type PaintSlots,
  parseSlots,
  serializeSlots,
  slotLabel,
} from "./paintbrush";
import {
  activeTheme,
  type BrandSettings,
  DEFAULT_SETTINGS,
  extractPaletteFromPixels,
  getActiveSettings,
  normalizeHex,
  parsePalette,
  readStoredSettings,
  serializeSettings,
  setActiveSettings,
} from "./settings";
import { makeGuard } from "./ui/guard";
import { describeError, installErrorReporting } from "./ui/report";
import { installTabs } from "./ui/tabs";
import { createToast } from "./ui/toast";
import { formatVersion } from "./ui/version";

// Every describeError call and the footer read this one formatted constant.
const APP_VERSION = formatVersion(__APP_VERSION__);

const STORAGE_KEY = "smt.brand.v1";
const PALETTE_SLOTS = [
  "primary",
  "accent",
  "input",
  "formula",
  "link",
  "external",
  "partial",
] as const;
type PaletteSlot = (typeof PALETTE_SLOTS)[number];

const getElement = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing element #${id}`);
  return element as T;
};

const connectionStatus = getElement<HTMLSpanElement>("connection-status");
const actionButtons = Array.from(
  document.querySelectorAll<HTMLButtonElement>("[data-action]"),
);
const toast = createToast(getElement("toast"));

function setBusy(busy: boolean): void {
  for (const button of actionButtons) button.disabled = busy;
}

function errorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  return "Excel could not complete that action.";
}

async function refreshSelection(): Promise<void> {
  try {
    const summary = await inspectSelection();
    // -1 means the selection was too large to read (whole column/row click).
    const metric = (value: number): string =>
      value < 0 ? "—" : value.toLocaleString();
    getElement("selection-address").textContent = summary.address;
    getElement("metric-cells").textContent = metric(summary.cells);
    getElement("metric-formulas").textContent = metric(summary.formulas);
    getElement("metric-errors").textContent = metric(summary.errors);
    getElement("metric-blanks").textContent = metric(summary.blanks);
  } catch (error) {
    toast.show(errorMessage(error), "error");
  }
}

// ---------------------------------------------------------------------------
// Audit overlay and Smart Track
// ---------------------------------------------------------------------------

const TRACE_STACK_LIMIT = 20;

interface TraceView {
  direction: TraceDirection;
  result: TraceResult;
}

const traceStack: TraceView[] = [];
let traceView: TraceView | null = null;
let auditOn = false;

function renderAuditState(): void {
  getElement("audit-state").textContent = auditOn ? "On" : "Off";
}

// The undo slot and the copy source are module state in excel.ts; the pane
// reads them back after every action so both rows say what they will do.
function renderActionState(): void {
  getElement("undo-target").textContent = undoTarget() ?? "Nothing to undo yet";
  const source = copySourceLabel();
  getElement("paste-source").textContent = source
    ? `Copy source: ${source}`
    : "Mark a source, then paste it into any selection.";
}

function renderTrace(): void {
  const panel = getElement<HTMLDivElement>("trace-panel");
  const chips = getElement<HTMLDivElement>("trace-chips");
  // Clearing before the early return keeps stale chips from firing when the
  // panel comes back.
  chips.replaceChildren();

  if (!traceView) {
    panel.hidden = true;
    return;
  }

  const { direction, result } = traceView;
  getElement("trace-origin").textContent = `${result.origin} · ${direction}`;
  getElement<HTMLButtonElement>("trace-back").disabled =
    traceStack.length === 0;

  if (result.areas.length === 0) {
    const empty = document.createElement("p");
    empty.className = "hint";
    empty.textContent = `No direct ${direction}.`;
    chips.append(empty);
  }

  for (const area of result.areas) {
    const label = `${area.sheet}!${area.address}`;
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "chip";
    chip.textContent = label;
    chip.title = `Select ${label} (${area.cellCount.toLocaleString()} cells)`;
    chip.addEventListener("click", () => void guard(() => walkTo(area)));
    chips.append(chip);
  }

  panel.hidden = false;
}

function traceMessage(direction: TraceDirection, count: number): string {
  if (count === 0) return `No direct ${direction}`;
  return `${count} direct ${count === 1 ? direction.slice(0, -1) : direction}`;
}

async function showTrace(
  direction: TraceDirection,
  jump: boolean,
): Promise<string> {
  const result = await traceActiveCell(direction);
  traceView = { direction, result };
  renderTrace();

  // A shortcut can fire with the pane closed, so the keystroke jumps instead.
  const first = result.areas[0];
  if (jump && first) await selectArea(first);

  return traceMessage(direction, result.areas.length);
}

// A fresh trace is a new walk, so the back stack starts empty.
function startTrace(direction: TraceDirection, jump: boolean): Promise<string> {
  traceStack.length = 0;
  return showTrace(direction, jump);
}

async function walkTo(area: TraceArea): Promise<string> {
  const current = traceView;
  if (!current) return "Nothing to trace";

  await selectArea(area);
  if (traceStack.length >= TRACE_STACK_LIMIT) traceStack.shift();
  traceStack.push(current);
  return showTrace(current.direction, false);
}

async function traceBack(): Promise<string> {
  const previous = traceStack.pop();
  if (!previous) return "Nothing to go back to";

  traceView = previous;
  renderTrace();
  await selectArea(parseAddress(previous.result.origin));
  return `Back at ${previous.result.origin}`;
}

async function toggleAudit(): Promise<string> {
  auditOn = await toggleAuditOverlay();
  renderAuditState();
  return auditOn ? "Audit overlay on" : "Audit overlay off";
}

// ---------------------------------------------------------------------------
// Paintbrush slots
// ---------------------------------------------------------------------------

// Three captured formats, kept on the machine like the brand palette: they
// outlive the pane, and nothing about them is written into the workbook.
const PAINT_KEY = "smt.paint.v1";
let paintSlots: PaintSlots = emptySlots();

function loadPaintSlots(): void {
  let raw: string | null = null;
  try {
    raw = localStorage.getItem(PAINT_KEY);
  } catch {
    raw = null;
  }
  paintSlots = parseSlots(raw);
}

function renderPaintSlots(): void {
  paintSlots.forEach((slot, index) => {
    getElement(`paint-slot-${index + 1}`).textContent = slotLabel(slot);
  });
}

async function capturePaintSlot(index: number): Promise<string> {
  const slot = await captureSlot(index);
  paintSlots[index - 1] = slot;
  try {
    localStorage.setItem(PAINT_KEY, serializeSlots(paintSlots));
  } catch {
    // Storage can be unavailable in private webviews; slots stay in memory.
  }
  renderPaintSlots();
  return `Slot ${index}: ${slotLabel(slot)}`;
}

async function applyPaintSlot(index: number): Promise<string> {
  await applySlot(index, paintSlots[index - 1] ?? null);
  return `Painted slot ${index}`;
}

async function dispatch(action: string): Promise<string> {
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
        await autocolorSelection();
        break;
      case "insert-color-key":
        await insertColorKey();
        break;
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
      default:
        throw new Error(`Unknown action: ${action}`);
    }
  }

  return "Selection updated";
}

// Every pane interaction runs through here: buttons off, toast on, busy cleared.
// renderActionState also runs after a failure: a capture may have replaced
// the undo slot already.
const guard = makeGuard({
  setBusy,
  notify: toast.show,
  describe: (error, action) =>
    describeError(error, { host: "Excel", version: APP_VERSION }, action),
  after: refreshSelection,
  decorate: (message) =>
    lastUndoSkipped() ? `${message} (too large for undo)` : message,
  finally: renderActionState,
});

// ---------------------------------------------------------------------------
// Ribbon commands and keyboard shortcuts (shared runtime)
// ---------------------------------------------------------------------------

interface CommandEvent {
  completed: () => void;
}

function registerCommands(): void {
  if (!Office.actions?.associate) return;

  const commands: Record<string, () => Promise<unknown>> = {
    SMT_AUTOCOLOR: autocolorSelection,
    SMT_AUDIT: toggleAudit,
    // With the pane closed there is nothing to read, so the keystroke jumps to
    // the first result instead.
    SMT_TRACE_PRE: () => startTrace("precedents", true),
    SMT_TRACE_DEP: () => startTrace("dependents", true),
    SMT_FILLRIGHT: () => fastFillAuto("right"),
    SMT_FILLDOWN: () => fastFillAuto("down"),
    SMT_IFERROR: toggleIfErrorGuard,
    SMT_SCALEUP: () => scaleSelection(1000),
    SMT_SCALEDOWN: () => scaleSelection(0.001),
    SMT_UNDO: undoLastAction,
    SMT_COPYSRC: markCopySource,
    SMT_PASTE_VALUES: () => pasteSpecial("values"),
    SMT_PASTE_FORMATS: () => pasteSpecial("formats"),
    SMT_PASTE_EXACT: pastePreserveFormulas,
    SMT_PASTE_TRANSPOSE: () => pasteSpecial("transpose"),
    SMT_CAGR: insertCagr,
    SMT_ROUND: insertConsistentRounding,
    SMT_SIGN: applySignFlip,
    SMT_DEC_MORE: () => applyDecimalStep(1),
    SMT_DEC_LESS: () => applyDecimalStep(-1),
    SMT_CYC_GENERAL: () => applyNumberCycle("general"),
    SMT_CYC_DATE: () => applyNumberCycle("date"),
    SMT_CYC_CURRENCY: () => applyNumberCycle("currency"),
    SMT_CYC_PERCENT: () => applyNumberCycle("percent"),
    SMT_CYC_MULTIPLE: () => applyNumberCycle("multiple"),
    SMT_CYC_TITLE: () => applyRowStyleCycle("title"),
    SMT_CYC_RESULT: () => applyRowStyleCycle("result"),
    SMT_CYC_ITEM: () => applyRowStyleCycle("item"),
    SMT_CYC_FILL: applyFillCycle,
    SMT_CYC_FONT: applyFontColorCycle,
    SMT_CYC_BORDER: applyBorderCycle,
    SMT_CYC_ROWH: applyRowHeightCycle,
    SMT_CYC_COLW: applyColumnWidthCycle,
    SMT_PAINT_CAP1: () => capturePaintSlot(1),
    SMT_PAINT_CAP2: () => capturePaintSlot(2),
    SMT_PAINT_CAP3: () => capturePaintSlot(3),
    SMT_PAINT_APP1: () => applyPaintSlot(1),
    SMT_PAINT_APP2: () => applyPaintSlot(2),
    SMT_PAINT_APP3: () => applyPaintSlot(3),
    SMT_WATERFALL: insertWaterfall,
    SMT_TORNADO: insertTornado,
    SMT_CHARTFMT: formatSelectedChart,
    SMT_CHART_CAGR: addCagrLabel,
    SMT_UNPIVOT: unpivotSelection,
    SMT_TOC: insertTocSheet,
    SMT_FIND: focusFind,
    SMT_STYLES_SCAN: focusStyles,
  };

  for (const [id, run] of Object.entries(commands)) {
    Office.actions.associate(id, (event?: CommandEvent) => {
      void run()
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

  Office.actions.associate("SMT_SHOWPANE", (event?: CommandEvent) => {
    void Promise.resolve(Office.addin?.showAsTaskpane())
      .catch(() => undefined)
      .finally(() => event?.completed());
  });
}

// ---------------------------------------------------------------------------
// Brand dashboard
// ---------------------------------------------------------------------------

// The workbook copy is best effort in both directions: a host that will not
// answer must not break the palette change the modeller just made, so failures
// land on the pane's error surface instead of in the caller.
function reportBrandStoreError(error: unknown, action: string): void {
  const { message, details } = describeError(
    error,
    { host: "Excel", version: APP_VERSION },
    action,
  );
  toast.show(message, "error", details);
}

function persistSettings(): void {
  const json = serializeSettings(getActiveSettings());
  try {
    localStorage.setItem(STORAGE_KEY, json);
  } catch {
    // Storage can be unavailable in private webviews; settings stay in memory.
  }
  // Saved with the file as well, so the palette follows the model to another
  // computer rather than living only on the machine that set it.
  // Only a connected workbook can carry the palette; the dev browser has none.
  if (!excelReady) return;
  void writeWorkbookBrand(json).catch((error: unknown) => {
    reportBrandStoreError(error, "save the brand to the workbook");
  });
}

function loadSettings(): void {
  let raw: string | null = null;
  try {
    raw = localStorage.getItem(STORAGE_KEY);
  } catch {
    raw = null;
  }
  setActiveSettings(readStoredSettings(raw));
}

// Boot, Excel only: a palette saved in the workbook wins over the machine
// default loadSettings already applied. No setting means a workbook that never
// carried a brand, which keeps that default rather than the shipped colors.
async function adoptWorkbookBrand(): Promise<void> {
  let json: string | null = null;
  try {
    json = await readWorkbookBrand();
  } catch (error) {
    reportBrandStoreError(error, "read the brand from the workbook");
    return;
  }

  const stored = json === null ? null : parsePalette(json);
  if (!stored) return;
  setActiveSettings(stored);
  renderBrand();
  syncAutocolorOnEdit();
}

function renderBrand(): void {
  const settings = getActiveSettings();
  const theme = activeTheme();

  for (const slot of PALETTE_SLOTS) {
    const picker = document.querySelector<HTMLInputElement>(
      `[data-slot-color="${slot}"]`,
    );
    const hex = document.querySelector<HTMLInputElement>(
      `[data-slot-hex="${slot}"]`,
    );
    if (picker) picker.value = settings[slot];
    if (hex) hex.value = settings[slot];
  }

  getElement<HTMLSelectElement>("setting-font").value = settings.font;
  getElement<HTMLSelectElement>("setting-currency").value = settings.currency;
  getElement<HTMLInputElement>("setting-autocolor-edit").checked =
    settings.autocolorOnEdit;
  getElement("currency-format-button").textContent =
    `${settings.currency ? `${settings.currency} ` : ""}1,234`;

  const preview = getElement<HTMLTableElement>("brand-preview");
  preview.style.fontFamily = `"${settings.font}", "Segoe UI", sans-serif`;
  preview.style.setProperty("--pv-title-fill", theme.titleFill);
  preview.style.setProperty("--pv-title-text", theme.titleText);
  preview.style.setProperty("--pv-header-fill", theme.headerFill);
  preview.style.setProperty("--pv-header-border", theme.headerBorder);
  preview.style.setProperty("--pv-result-fill", theme.resultFill);
  preview.style.setProperty("--pv-result-border", theme.resultBorder);
  preview.style.setProperty("--pv-input", theme.inputFont);
  preview.style.setProperty("--pv-formula", theme.formulaFont);
  preview.style.setProperty("--pv-link", theme.linkFont);
  preview.style.setProperty("--pv-external", theme.externalFont);
  preview.style.setProperty("--pv-partial", theme.partialFont);
}

// Excel is only there when the pane runs inside the host; the toast reports the rest.
function syncAutocolorOnEdit(): void {
  setAutocolorOnEdit(getActiveSettings().autocolorOnEdit).catch((error) => {
    toast.show(errorMessage(error), "error");
  });
}

function applySettings(next: BrandSettings, message?: string): void {
  setActiveSettings(next);
  persistSettings();
  renderBrand();
  syncAutocolorOnEdit();
  if (message) toast.show(message);
}

function updateSetting(patch: Partial<BrandSettings>, message?: string): void {
  applySettings({ ...getActiveSettings(), ...patch }, message);
}

let logoAssignIndex = 0;

function renderLogoSwatches(colors: string[]): void {
  const strip = getElement<HTMLDivElement>("logo-swatches");
  const hint = getElement<HTMLParagraphElement>("logo-hint");
  strip.replaceChildren();
  logoAssignIndex = 0;

  if (colors.length === 0) {
    strip.hidden = true;
    hint.hidden = true;
    toast.show("No usable colors found in that image.", "error");
    return;
  }

  for (const color of colors) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "logo-swatch";
    button.style.background = color;
    button.title = color;
    button.setAttribute("aria-label", `Use ${color}`);
    button.addEventListener("click", () => {
      const slot: PaletteSlot =
        logoAssignIndex % 2 === 0 ? "primary" : "accent";
      logoAssignIndex += 1;
      const label = slot === "primary" ? "Primary" : "Accent";
      updateSetting({ [slot]: color }, `${label} set to ${color}`);
    });
    strip.append(button);
  }

  strip.hidden = false;
  hint.hidden = false;
}

function extractLogoColors(file: File): void {
  const url = URL.createObjectURL(file);
  const image = new Image();
  image.onload = () => {
    URL.revokeObjectURL(url);
    const scale = Math.min(1, 64 / Math.max(image.width, image.height));
    const width = Math.max(1, Math.round(image.width * scale));
    const height = Math.max(1, Math.round(image.height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) {
      toast.show("Could not read that image.", "error");
      return;
    }
    context.drawImage(image, 0, 0, width, height);
    const { data } = context.getImageData(0, 0, width, height);
    renderLogoSwatches(extractPaletteFromPixels(data, 6));
  };
  image.onerror = () => {
    URL.revokeObjectURL(url);
    toast.show("That file is not a readable image.", "error");
  };
  image.src = url;
}

async function copyPaletteJson(): Promise<void> {
  const json = serializeSettings(getActiveSettings());
  try {
    await navigator.clipboard.writeText(json);
    toast.show("Palette JSON copied");
  } catch {
    const area = document.createElement("textarea");
    area.value = json;
    document.body.append(area);
    area.select();
    const copied = document.execCommand("copy");
    area.remove();
    toast.show(
      copied ? "Palette JSON copied" : "Copy failed",
      copied ? "success" : "error",
    );
  }
}

function wireBrand(): void {
  for (const slot of PALETTE_SLOTS) {
    const picker = document.querySelector<HTMLInputElement>(
      `[data-slot-color="${slot}"]`,
    );
    const hex = document.querySelector<HTMLInputElement>(
      `[data-slot-hex="${slot}"]`,
    );
    picker?.addEventListener("input", () => {
      const value = normalizeHex(picker.value);
      if (value) updateSetting({ [slot]: value });
    });
    hex?.addEventListener("change", () => {
      const value = normalizeHex(hex.value);
      if (!value) {
        toast.show(`"${hex.value}" is not a hex color like #B27E54.`, "error");
        renderBrand();
        return;
      }
      updateSetting({ [slot]: value });
    });
  }

  getElement<HTMLSelectElement>("setting-font").addEventListener(
    "change",
    (event) => {
      updateSetting({ font: (event.target as HTMLSelectElement).value });
    },
  );
  getElement<HTMLSelectElement>("setting-currency").addEventListener(
    "change",
    (event) => {
      updateSetting({ currency: (event.target as HTMLSelectElement).value });
    },
  );

  getElement<HTMLInputElement>("setting-autocolor-edit").addEventListener(
    "change",
    (event) => {
      const on = (event.target as HTMLInputElement).checked;
      updateSetting(
        { autocolorOnEdit: on },
        on ? "Autocolor runs on every edit" : "Autocolor on edit is off",
      );
    },
  );

  getElement<HTMLButtonElement>("reset-brand").addEventListener("click", () => {
    applySettings({ ...DEFAULT_SETTINGS }, "Palette reset to house defaults");
    const strip = getElement<HTMLDivElement>("logo-swatches");
    strip.replaceChildren();
    strip.hidden = true;
    getElement<HTMLParagraphElement>("logo-hint").hidden = true;
  });

  getElement<HTMLInputElement>("logo-file").addEventListener(
    "change",
    (event) => {
      const input = event.target as HTMLInputElement;
      const file = input.files?.[0];
      if (file) extractLogoColors(file);
      input.value = "";
    },
  );

  getElement<HTMLButtonElement>("export-brand").addEventListener(
    "click",
    () => {
      void copyPaletteJson();
    },
  );

  getElement<HTMLInputElement>("import-file").addEventListener(
    "change",
    async (event) => {
      const input = event.target as HTMLInputElement;
      const file = input.files?.[0];
      input.value = "";
      if (!file) return;
      const parsed = parsePalette(await file.text());
      if (!parsed) {
        toast.show("That file is not a valid palette JSON.", "error");
        return;
      }
      applySettings(parsed, "Palette imported");
    },
  );
}

// ---------------------------------------------------------------------------
// Workbook tools
// ---------------------------------------------------------------------------

const DELETE_CONFIRM_MS = 5_000;

let excelReady = false;
let brokenList: string[] = [];
let deleteArmed = false;
let deleteTimer: number | undefined;
// Null until the first scan: the delete acts on a list, never on a guess.
let styleScan: StyleScan | null = null;
let stylesArmed = false;
let stylesTimer: number | undefined;

async function goToSheet(name: string): Promise<string> {
  await activateSheet(name);
  await refreshSheets();
  return `Switched to ${name}`;
}

async function toggleSheet(name: string, visible: boolean): Promise<string> {
  await setSheetVisibility(name, visible);
  await refreshSheets();
  return visible ? `${name} is visible again` : `${name} is hidden`;
}

function sheetRow(sheet: SheetEntry): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "sheet-row";
  const visible = sheet.visibility === "Visible";
  const locked = sheet.visibility === "VeryHidden";
  if (locked) row.classList.add("locked");
  if (sheet.active) {
    row.classList.add("active");
    row.setAttribute("aria-current", "true");
  }

  // Excel refuses to activate a sheet nobody can see, so only a visible name
  // is a button.
  const name = document.createElement(visible ? "button" : "span");
  name.className = visible ? "sheet-name" : "sheet-name muted";
  name.textContent = sheet.name;
  if (name instanceof HTMLButtonElement) {
    name.type = "button";
    name.title = `Go to ${sheet.name}`;
    name.addEventListener(
      "click",
      () => void guard(() => goToSheet(sheet.name)),
    );
  }
  row.append(name);

  const badge = document.createElement("span");
  badge.className = visible ? "sheet-badge" : "sheet-badge off";
  badge.textContent = locked ? "Very hidden" : visible ? "Visible" : "Hidden";
  row.append(badge);

  if (locked) {
    // No toggle: very hidden is set outside Excel's UI and stays that way. The
    // empty slot keeps the row the same height as the ones that have a button.
    const slot = document.createElement("span");
    slot.className = "eye-slot";
    row.append(slot);
    return row;
  }

  const eye = document.createElement("button");
  eye.type = "button";
  eye.className = "eye";
  eye.textContent = visible ? "◉" : "○";
  eye.title = visible ? `Hide ${sheet.name}` : `Show ${sheet.name}`;
  eye.setAttribute("aria-label", eye.title);
  eye.addEventListener(
    "click",
    () => void guard(() => toggleSheet(sheet.name, !visible)),
  );
  row.append(eye);
  return row;
}

async function refreshSheets(): Promise<void> {
  const list = getElement<HTMLDivElement>("sheet-list");
  // Rows close over the sheet they were built from; drop them before rebuilding.
  list.replaceChildren();

  if (!excelReady) {
    const hint = document.createElement("p");
    hint.className = "hint";
    hint.textContent = "Connect to Excel to list the sheets.";
    list.append(hint);
    return;
  }

  try {
    const sheets = await listSheets();
    for (const sheet of sheets) list.append(sheetRow(sheet));
  } catch (error) {
    toast.show(errorMessage(error), "error");
  }
}

function brokenCount(): string {
  return `${brokenList.length} broken ${brokenList.length === 1 ? "name" : "names"}`;
}

function deleteLabel(): string {
  return `Delete ${brokenCount()}`;
}

function disarmDelete(): void {
  window.clearTimeout(deleteTimer);
  deleteArmed = false;
  const button = getElement<HTMLButtonElement>("delete-names");
  button.classList.remove("armed");
  button.textContent = deleteLabel();
}

// Deleting a name cannot be undone by us or by Excel, so the first click only
// arms the button and the arming lapses on its own.
function armDelete(): void {
  const button = getElement<HTMLButtonElement>("delete-names");
  deleteArmed = true;
  button.classList.add("armed");
  button.textContent = "Click again to confirm";
  deleteTimer = window.setTimeout(disarmDelete, DELETE_CONFIRM_MS);
}

function renderNames(scanned: boolean): void {
  const result = getElement("names-result");
  const listed = brokenList.slice(0, 6).join(", ");
  const rest = brokenList.length > 6 ? ", …" : "";

  if (!scanned) result.textContent = "Not scanned yet.";
  else if (brokenList.length === 0) result.textContent = "No broken names.";
  else result.textContent = `${brokenCount()}: ${listed}${rest}`;

  disarmDelete();
  getElement<HTMLButtonElement>("delete-names").hidden =
    brokenList.length === 0;
}

async function scanNames(): Promise<string> {
  brokenList = await listBrokenNames();
  renderNames(true);
  return brokenList.length === 0 ? "No broken names" : brokenCount();
}

async function deleteNames(): Promise<string> {
  const removed = await deleteBrokenNames();
  brokenList = [];
  renderNames(true);
  return `Deleted ${removed} broken ${removed === 1 ? "name" : "names"}`;
}

async function insertTocSheet(): Promise<string> {
  await insertToc();
  await refreshSheets();
  return "Contents sheet updated";
}

// ---------------------------------------------------------------------------
// Super Find
// ---------------------------------------------------------------------------

const FIND_TEXT_LIMIT = 90;
const FIND_ICONS: Record<FindHit["kind"], string> = {
  cell: "▤",
  name: "⌗",
  sheet: "☰",
};

function hitLabel(hit: FindHit): string {
  if (hit.kind === "name") return `Name · ${hit.address}`;
  return `${hit.sheet}!${hit.address}`;
}

// A long label or a formula would push the row out of the pane; the whole text
// stays in the tooltip.
function clipText(text: string): string {
  if (text.length <= FIND_TEXT_LIMIT) return text;
  return `${text.slice(0, FIND_TEXT_LIMIT)}…`;
}

async function jumpTo(hit: FindHit): Promise<string> {
  await jumpToHit(hit);
  // The jump moved the workbook, so the explorer above marks the sheet it
  // landed on rather than the one it left.
  await refreshSheets();
  return `Jumped to ${hitLabel(hit)}`;
}

function findRow(hit: FindHit): HTMLButtonElement {
  const row = document.createElement("button");
  row.type = "button";
  row.title = `Go to ${hitLabel(hit)}: ${hit.text}`;

  const icon = document.createElement("span");
  icon.className = "action-icon names";
  icon.textContent = FIND_ICONS[hit.kind];

  const where = document.createElement("strong");
  where.textContent = hitLabel(hit);
  const what = document.createElement("small");
  what.textContent = clipText(hit.text);
  const text = document.createElement("span");
  text.append(where, what);

  row.append(icon, text);
  row.addEventListener("click", () => void guard(() => jumpTo(hit)));
  return row;
}

function findSummary(result: FindResult): string {
  const count = result.hits.length;
  const capped = count >= FIND_HIT_CAP ? ` (first ${FIND_HIT_CAP})` : "";
  const skipped =
    result.skippedSheets.length > 0
      ? ` Too large to search: ${result.skippedSheets.join(", ")}.`
      : "";
  if (count === 0) return `No matches.${skipped}`;
  return `${count} ${count === 1 ? "hit" : "hits"}${capped}.${skipped}`;
}

// Rows close over the hit they jump to; drop them before rebuilding. A null
// result is the state before the first search.
function renderFind(result: FindResult | null): void {
  const list = getElement<HTMLDivElement>("find-results");
  list.replaceChildren();
  // An empty list still costs a row gap under the button, so it goes away.
  list.hidden = !result || result.hits.length === 0;

  if (!result) {
    getElement("find-hint").textContent =
      "Searches values, defined names and sheet names on every sheet.";
    return;
  }
  for (const hit of result.hits) list.append(findRow(hit));
  getElement("find-hint").textContent = findSummary(result);
}

async function runFind(): Promise<string> {
  const query = getElement<HTMLInputElement>("find-query").value.trim();
  if (query === "") {
    renderFind(null);
    throw new Error("Type something to find first.");
  }

  const result = await findInWorkbook(query, {
    matchCase: getElement<HTMLInputElement>("find-case").checked,
    inFormulas: getElement<HTMLInputElement>("find-formulas").checked,
  });
  renderFind(result);
  return findSummary(result);
}

// The results live in the pane, so the shortcut opens it and puts the caret in
// the box rather than repeating a query the modeller cannot see.
async function focusFind(): Promise<string> {
  await Promise.resolve(Office.addin?.showAsTaskpane()).catch(() => undefined);
  tabs.activate("tab-workbook");
  await refreshSheets();

  const input = getElement<HTMLInputElement>("find-query");
  input.focus();
  input.select();
  return "Find ready";
}

// ---------------------------------------------------------------------------
// Style scrubber
// ---------------------------------------------------------------------------

function unusedCount(): string {
  const count = styleScan?.unused.length ?? 0;
  return `${count} unused ${count === 1 ? "style" : "styles"}`;
}

function stylesDeleteLabel(): string {
  return `Delete ${unusedCount()}`;
}

function disarmStyles(): void {
  window.clearTimeout(stylesTimer);
  stylesArmed = false;
  const button = getElement<HTMLButtonElement>("styles-delete");
  button.classList.remove("armed");
  button.textContent = stylesDeleteLabel();
}

// Deleting a style restyles every cell wearing it and no undo brings it back,
// so the first click only arms the button and the arming lapses on its own.
function armStyles(): void {
  const button = getElement<HTMLButtonElement>("styles-delete");
  stylesArmed = true;
  button.classList.add("armed");
  button.textContent = "Click again to confirm";
  stylesTimer = window.setTimeout(disarmStyles, DELETE_CONFIRM_MS);
}

// A sheet too large to read could be wearing any of these styles, so the count
// is stated as incomplete and the delete stays out of reach until it is not.
function stylesSummary(scan: StyleScan): string {
  const table = ` of ${scan.total} in the workbook`;
  const skipped =
    scan.skippedSheets.length > 0
      ? ` Some sheets were too large to scan: ${scan.skippedSheets.join(", ")}.`
      : "";
  if (scan.unused.length === 0) {
    return `No unused custom styles${table}.${skipped}`;
  }
  return `${unusedCount()}${table}.${skipped}`;
}

function styleRow(name: string): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "sheet-row";
  // Muted: a style name is a label, not somewhere to click through to.
  const label = document.createElement("span");
  label.className = "sheet-name muted";
  label.textContent = name;
  row.append(label);
  return row;
}

function renderStyles(): void {
  const list = getElement<HTMLDivElement>("styles-list");
  const button = getElement<HTMLButtonElement>("styles-delete");
  list.replaceChildren();

  if (!styleScan) {
    getElement("styles-result").textContent = "Not scanned yet.";
    list.hidden = true;
    button.hidden = true;
    return;
  }

  for (const name of styleScan.unused) list.append(styleRow(name));
  list.hidden = styleScan.unused.length === 0;
  getElement("styles-result").textContent = stylesSummary(styleScan);

  disarmStyles();
  button.hidden = styleScan.unused.length === 0;
  button.disabled = styleScan.skippedSheets.length > 0;
}

async function scanStyles(): Promise<string> {
  styleScan = await listUnusedStyles();
  renderStyles();
  return stylesSummary(styleScan);
}

// Rescanned afterwards rather than assumed: the style table is what shrank, and
// the pane says what is left of it.
async function deleteStyles(): Promise<string> {
  const removed = await deleteUnusedStyles(styleScan?.unused ?? []);
  styleScan = await listUnusedStyles();
  renderStyles();
  return `Deleted ${removed} unused ${removed === 1 ? "style" : "styles"}`;
}

// The list lives in the pane, so the command opens it on the Workbook tab and
// leaves the result on screen rather than reporting a number and forgetting it.
async function focusStyles(): Promise<string> {
  await Promise.resolve(Office.addin?.showAsTaskpane()).catch(() => undefined);
  tabs.activate("tab-workbook");
  return scanStyles();
}

getElement("app-version").textContent = APP_VERSION;

// Installed first so a throw during the rest of boot is still reported.
installErrorReporting(
  { host: "Excel", version: APP_VERSION },
  (message, details) => toast.show(message, "error", details),
);
loadSettings();
loadPaintSlots();
const tabs = installTabs(getElement("tab-bar"));
getElement<HTMLButtonElement>("tab-workbook").addEventListener(
  "click",
  // Sheets change without the pane hearing about it, so the explorer is
  // read when it comes into view rather than on every selection change.
  () => void refreshSheets(),
);
wireBrand();
renderBrand();
renderPaintSlots();
renderAuditState();
renderActionState();
renderNames(false);
renderFind(null);
renderStyles();

Office.onReady(async ({ host }) => {
  if (host === Office.HostType.PowerPoint) {
    location.replace("pptpane.html");
    return;
  }

  if (host !== Office.HostType.Excel) {
    connectionStatus.textContent = "Excel required";
    connectionStatus.className = "connection error";
    return;
  }

  // The manifest no longer guarantees ExcelApi 1.9 at the top level now that
  // PowerPoint is a second host, so the pane checks for itself.
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    connectionStatus.textContent = "Excel 2021 / Microsoft 365 required";
    connectionStatus.className = "connection error";
    return;
  }

  connectionStatus.textContent = "Excel connected";
  connectionStatus.className = "connection ready";
  excelReady = true;

  registerCommands();
  syncAutocolorOnEdit();
  void adoptWorkbookBrand();

  installLinksTab({
    guard,
    toast,
    relay: new RelayClient(relayBaseUrl(document.baseURI)),
    keyStore: officeKeyStore(),
    root: document,
  });

  for (const button of actionButtons) {
    button.addEventListener("click", () => {
      const action = button.dataset.action;
      if (action) void guard(() => dispatch(action), action);
    });
  }

  getElement<HTMLButtonElement>("refresh-selection").addEventListener(
    "click",
    () => void refreshSelection(),
  );

  getElement<HTMLButtonElement>("trace-back").addEventListener(
    "click",
    () => void guard(traceBack),
  );

  getElement<HTMLButtonElement>("refresh-sheets").addEventListener(
    "click",
    () => void refreshSheets(),
  );

  // Enter is what a search box owes the hands already on the keyboard.
  getElement<HTMLInputElement>("find-query").addEventListener(
    "keydown",
    (event) => {
      if (event.key === "Enter") void guard(runFind, "find");
    },
  );

  getElement<HTMLButtonElement>("delete-names").addEventListener(
    "click",
    () => {
      if (!deleteArmed) {
        armDelete();
        return;
      }
      disarmDelete();
      void guard(deleteNames);
    },
  );

  getElement<HTMLButtonElement>("styles-delete").addEventListener(
    "click",
    () => {
      if (!stylesArmed) {
        armStyles();
        return;
      }
      disarmStyles();
      void guard(() => dispatch("styles-delete"), "styles-delete");
    },
  );

  // Debounced: dragging a selection fires the event continuously.
  let selectionTimer: number | undefined;
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => {
      window.clearTimeout(selectionTimer);
      selectionTimer = window.setTimeout(() => void refreshSelection(), 150);
    },
  );

  void restorePersistedOverlay()
    .then((restored) => {
      if (restored) {
        toast.show("Audit overlay fills from the last session were restored.");
      }
    })
    .catch(() => undefined);

  await refreshSelection();
});
