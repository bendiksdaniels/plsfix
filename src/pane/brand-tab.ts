// The brand dashboard: the palette, font, currency and language settings,
// persisted to localStorage and to the workbook (the workbook copy wins at
// boot over the machine one), plus the logo color picker and the two JSON
// import/export controls ./brand-io backs. Office.js only through ../excel.

import {
  type ExcelSeparators,
  readWorkbookBrand,
  setAutocolorOnEdit,
  writeWorkbookBrand,
} from "../excel";
import {
  currencyFormat,
  formatAmount,
  isLanguage,
  LANGUAGES,
  numberStyle,
  separatorSample,
  separatorsMatch,
} from "../numbers";
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
} from "../settings";
import { copyPaletteJson, readPaletteFile } from "./brand-io";
import { getElement } from "../ui/dom";
import { describeError } from "../ui/report";
import { APP_VERSION, errorMessage, isExcelReady, toast } from "./shared";

const STORAGE_KEY = "plsfix.brand.v1";
// Excel's own separators, read once the host is there; null until then or below ExcelApi 1.11.
let excelSeparators: ExcelSeparators | null = null;
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

// Boot calls this once readSeparators resolves: the note redraws with
// whatever Excel reports rather than staying hidden past that point.
export function applyExcelSeparators(found: ExcelSeparators | null): void {
  excelSeparators = found;
  renderSeparatorsNote(getActiveSettings());
}

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
  if (!isExcelReady()) return;
  void writeWorkbookBrand(json).catch((error: unknown) => {
    reportBrandStoreError(error, "save the brand to the workbook");
  });
}

export function loadSettings(): void {
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
export async function adoptWorkbookBrand(): Promise<void> {
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

export function renderBrand(): void {
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
  getElement<HTMLSelectElement>("setting-language").value = settings.language;
  getElement<HTMLSelectElement>("setting-currency").value = settings.currency;
  getElement<HTMLInputElement>("setting-autocolor-edit").checked =
    settings.autocolorOnEdit;
  getElement("currency-format-button").textContent = currencyFormat(
    settings.currency,
    settings.language,
    formatAmount(1234, settings.language),
  );
  renderSeparatorsNote(settings);

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

// Format codes stay #,##0: what they show is Excel's own separator setting, so
// the note says what Excel shows and, when that is not the house style, where
// to change it.
function renderSeparatorsNote(settings: BrandSettings): void {
  const note = getElement<HTMLParagraphElement>("separators-note");
  if (excelSeparators === null) {
    note.hidden = true;
    return;
  }
  const { decimal, thousands } = excelSeparators;
  const shown = separatorSample(decimal, thousands);
  if (separatorsMatch(decimal, thousands, settings.language)) {
    note.textContent = `Excel shows ${shown}, the house style.`;
  } else {
    const style = numberStyle(settings.language);
    const language =
      LANGUAGES.find((option) => option.code === settings.language)?.label ??
      settings.language;
    note.textContent = `Excel shows ${shown}; the ${language} style is ${separatorSample(style.decimal, style.grouping)}. Separators are an Excel setting: Excel > Preferences > Edit on Mac, File > Options > Advanced on Windows, thousands "${style.grouping}" and decimal ".".`;
  }
  note.hidden = false;
}

// There is no workbook to hang the edit handler on until boot confirms one, so
// a palette change before that must not become a red toast about a missing
// host. Boot calls this itself once connected, so nothing is lost.
export function syncAutocolorOnEdit(): void {
  if (!isExcelReady()) return;
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

export function wireBrand(): void {
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
        toast.show(`"${hex.value}" is not a hex color like #2EC4B6.`, "error");
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
  getElement<HTMLSelectElement>("setting-language").addEventListener(
    "change",
    (event) => {
      const value = (event.target as HTMLSelectElement).value;
      if (isLanguage(value)) updateSetting({ language: value });
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
    applySettings({ ...DEFAULT_SETTINGS }, "Palette reset to pls,fix defaults");
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
      // A file the webview cannot open is one sentence, not a rejection.
      const json = await file.text().catch(() => null);
      if (json === null) {
        toast.show("That file could not be read.", "error");
        return;
      }
      const parsed = readPaletteFile(json);
      if (!parsed) {
        toast.show("That file is not a valid palette JSON.", "error");
        return;
      }
      applySettings(parsed, "Palette imported");
    },
  );
}
