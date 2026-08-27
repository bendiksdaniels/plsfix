import "./styles.css";
import type { NumberCycleFamily, RowStyleKind } from "./cycles";
import {
  addCagrLabel,
  applyDecimalStep,
  applyFillCycle,
  applyFontColorCycle,
  applyNumberCycle,
  applyNumberFormat,
  applyPreset,
  applyRowStyleCycle,
  applySignFlip,
  autocolorSelection,
  clearFormats,
  copySourceLabel,
  fastFillAuto,
  formatSelectedChart,
  insertCagr,
  insertColorKey,
  insertWaterfall,
  inspectSelection,
  markCopySource,
  parseAddress,
  pasteSpecial,
  pastePreserveFormulas,
  scaleSelection,
  selectArea,
  setAutocolorOnEdit,
  toggleAuditOverlay,
  toggleIfErrorGuard,
  traceActiveCell,
  undoLastAction,
  undoTarget,
  type NumberFormatName,
  type PresetName,
  type TraceArea,
  type TraceDirection,
  type TraceResult,
} from "./excel";
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
const toast = getElement<HTMLDivElement>("toast");
const actionButtons = Array.from(
  document.querySelectorAll<HTMLButtonElement>("[data-action]"),
);
let toastTimer: number | undefined;

function showToast(message: string, kind: "success" | "error" = "success"): void {
  window.clearTimeout(toastTimer);
  toast.textContent = message;
  toast.className = `toast visible ${kind}`;
  toastTimer = window.setTimeout(() => {
    toast.className = "toast";
  }, 3200);
}

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
    getElement("selection-address").textContent = summary.address;
    getElement("metric-cells").textContent = summary.cells.toLocaleString();
    getElement("metric-formulas").textContent = summary.formulas.toLocaleString();
    getElement("metric-errors").textContent = summary.errors.toLocaleString();
    getElement("metric-blanks").textContent = summary.blanks.toLocaleString();
  } catch (error) {
    showToast(errorMessage(error), "error");
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
  getElement("undo-target").textContent =
    undoTarget() ?? "Nothing to undo yet";
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
  getElement<HTMLButtonElement>("trace-back").disabled = traceStack.length === 0;

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

async function dispatch(action: string): Promise<string> {
  if (action.startsWith("style-")) {
    await applyPreset(action.replace("style-", "") as PresetName);
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
      case "chart-format":
        await formatSelectedChart();
        return "Chart restyled to your brand";
      case "chart-cagr":
        return addCagrLabel();
      case "audit-toggle":
        return toggleAudit();
      case "trace-precedents":
      case "trace-dependents":
        return startTrace(action.replace("trace-", "") as TraceDirection, false);
      default:
        throw new Error(`Unknown action: ${action}`);
    }
  }

  return "Selection updated";
}

// Every pane interaction runs through here: buttons off, toast on, busy cleared.
async function guard(run: () => Promise<string>): Promise<void> {
  setBusy(true);
  try {
    const message = await run();
    await refreshSelection();
    showToast(message);
  } catch (error) {
    showToast(errorMessage(error), "error");
  } finally {
    // Also after a failure: a capture may have replaced the undo slot already.
    renderActionState();
    setBusy(false);
  }
}

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
    SMT_WATERFALL: insertWaterfall,
    SMT_CHARTFMT: formatSelectedChart,
    SMT_CHART_CAGR: addCagrLabel,
  };

  for (const [id, run] of Object.entries(commands)) {
    Office.actions.associate(id, (event?: CommandEvent) => {
      void run()
        .then(() => refreshSelection())
        .catch((error) => showToast(errorMessage(error), "error"))
        .finally(() => {
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

function persistSettings(): void {
  try {
    localStorage.setItem(STORAGE_KEY, serializeSettings(getActiveSettings()));
  } catch {
    // Storage can be unavailable in private webviews; settings stay in memory.
  }
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
    showToast(errorMessage(error), "error");
  });
}

function applySettings(next: BrandSettings, message?: string): void {
  setActiveSettings(next);
  persistSettings();
  renderBrand();
  syncAutocolorOnEdit();
  if (message) showToast(message);
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
    showToast("No usable colors found in that image.", "error");
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
      const slot: PaletteSlot = logoAssignIndex % 2 === 0 ? "primary" : "accent";
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
      showToast("Could not read that image.", "error");
      return;
    }
    context.drawImage(image, 0, 0, width, height);
    const { data } = context.getImageData(0, 0, width, height);
    renderLogoSwatches(extractPaletteFromPixels(data, 6));
  };
  image.onerror = () => {
    URL.revokeObjectURL(url);
    showToast("That file is not a readable image.", "error");
  };
  image.src = url;
}

async function copyPaletteJson(): Promise<void> {
  const json = serializeSettings(getActiveSettings());
  try {
    await navigator.clipboard.writeText(json);
    showToast("Palette JSON copied");
  } catch {
    const area = document.createElement("textarea");
    area.value = json;
    document.body.append(area);
    area.select();
    const copied = document.execCommand("copy");
    area.remove();
    showToast(copied ? "Palette JSON copied" : "Copy failed", copied ? "success" : "error");
  }
}

function wireTabs(): void {
  const tabs = [
    { tab: getElement<HTMLButtonElement>("tab-tools"), view: getElement("view-tools") },
    { tab: getElement<HTMLButtonElement>("tab-brand"), view: getElement("view-brand") },
  ];
  for (const current of tabs) {
    current.tab.addEventListener("click", () => {
      for (const other of tabs) {
        const active = other === current;
        other.tab.classList.toggle("active", active);
        other.tab.setAttribute("aria-selected", String(active));
        other.view.hidden = !active;
      }
    });
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
        showToast(`"${hex.value}" is not a hex color like #B27E54.`, "error");
        renderBrand();
        return;
      }
      updateSetting({ [slot]: value });
    });
  }

  getElement<HTMLSelectElement>("setting-font").addEventListener("change", (event) => {
    updateSetting({ font: (event.target as HTMLSelectElement).value });
  });
  getElement<HTMLSelectElement>("setting-currency").addEventListener("change", (event) => {
    updateSetting({ currency: (event.target as HTMLSelectElement).value });
  });

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

  getElement<HTMLInputElement>("logo-file").addEventListener("change", (event) => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) extractLogoColors(file);
    input.value = "";
  });

  getElement<HTMLButtonElement>("export-brand").addEventListener("click", () => {
    void copyPaletteJson();
  });

  getElement<HTMLInputElement>("import-file").addEventListener("change", async (event) => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    if (!file) return;
    const parsed = parsePalette(await file.text());
    if (!parsed) {
      showToast("That file is not a valid palette JSON.", "error");
      return;
    }
    applySettings(parsed, "Palette imported");
  });
}

loadSettings();
wireTabs();
wireBrand();
renderBrand();
renderAuditState();
renderActionState();

Office.onReady(async ({ host }) => {
  if (host !== Office.HostType.Excel) {
    connectionStatus.textContent = "Excel required";
    connectionStatus.className = "connection error";
    return;
  }

  connectionStatus.textContent = "Excel connected";
  connectionStatus.className = "connection ready";

  registerCommands();
  syncAutocolorOnEdit();

  for (const button of actionButtons) {
    button.addEventListener("click", () => {
      const action = button.dataset.action;
      if (action) void guard(() => dispatch(action));
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

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => void refreshSelection(),
  );

  await refreshSelection();
});
