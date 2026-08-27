import "./styles.css";
import {
  addIfError,
  applyNumberFormat,
  applyPreset,
  autocolorSelection,
  clearFormats,
  fastFill,
  inspectSelection,
  scaleSelection,
  type NumberFormatName,
  type PresetName,
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
const PALETTE_SLOTS = ["primary", "accent", "input", "formula", "link"] as const;
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

async function runAction(action: string): Promise<void> {
  setBusy(true);
  try {
    if (action.startsWith("style-")) {
      await applyPreset(action.replace("style-", "") as PresetName);
    } else if (action.startsWith("number-")) {
      await applyNumberFormat(action.replace("number-", "") as NumberFormatName);
    } else {
      switch (action) {
        case "clear-formats":
          await clearFormats();
          break;
        case "fill-right":
          await fastFill("right");
          break;
        case "fill-down":
          await fastFill("down");
          break;
        case "if-error":
          await addIfError();
          break;
        case "autocolor":
          await autocolorSelection();
          break;
        case "divide-1000":
          await scaleSelection(0.001);
          break;
        case "multiply-1000":
          await scaleSelection(1000);
          break;
        default:
          throw new Error(`Unknown action: ${action}`);
      }
    }

    await refreshSelection();
    showToast("Selection updated");
  } catch (error) {
    showToast(errorMessage(error), "error");
  } finally {
    setBusy(false);
  }
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
}

function applySettings(next: BrandSettings, message?: string): void {
  setActiveSettings(next);
  persistSettings();
  renderBrand();
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

  getElement<HTMLButtonElement>("reset-brand").addEventListener("click", () => {
    applySettings({ ...DEFAULT_SETTINGS }, "Palette reset to house defaults");
    getElement<HTMLDivElement>("logo-swatches").hidden = true;
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

Office.onReady(async ({ host }) => {
  if (host !== Office.HostType.Excel) {
    connectionStatus.textContent = "Excel required";
    connectionStatus.className = "connection error";
    return;
  }

  connectionStatus.textContent = "Excel connected";
  connectionStatus.className = "connection ready";

  for (const button of actionButtons) {
    button.addEventListener("click", () => {
      const action = button.dataset.action;
      if (action) void runAction(action);
    });
  }

  getElement<HTMLButtonElement>("refresh-selection").addEventListener(
    "click",
    () => void refreshSelection(),
  );

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => void refreshSelection(),
  );

  await refreshSelection();
});
