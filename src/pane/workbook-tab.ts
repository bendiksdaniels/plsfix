// Workbook tools: the sheet explorer (go to, show/hide), the broken-name
// scrubber and the contents-sheet writer. The "click again to confirm" arm on
// the delete button lapses on its own. Office.js only reaches here through
// ../excel.

import {
  activateSheet,
  deleteBrokenNames,
  insertToc,
  listBrokenNames,
  listSheets,
  setSheetVisibility,
  type SheetEntry,
} from "../excel";
import { getElement } from "../ui/dom";
import {
  DELETE_CONFIRM_MS,
  errorMessage,
  guard,
  isExcelReady,
  toast,
} from "./shared";

let brokenList: string[] = [];
let deleteArmed = false;
let deleteTimer: number | undefined;

export function isDeleteArmed(): boolean {
  return deleteArmed;
}

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

export async function refreshSheets(): Promise<void> {
  const list = getElement<HTMLDivElement>("sheet-list");
  // Rows close over the sheet they were built from; drop them before rebuilding.
  list.replaceChildren();

  if (!isExcelReady()) {
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

export function disarmDelete(): void {
  window.clearTimeout(deleteTimer);
  deleteArmed = false;
  const button = getElement<HTMLButtonElement>("delete-names");
  button.classList.remove("armed");
  button.textContent = deleteLabel();
}

// Deleting a name cannot be undone by us or by Excel, so the first click only
// arms the button and the arming lapses on its own.
export function armDelete(): void {
  const button = getElement<HTMLButtonElement>("delete-names");
  deleteArmed = true;
  button.classList.add("armed");
  button.textContent = "Click again to confirm";
  deleteTimer = window.setTimeout(disarmDelete, DELETE_CONFIRM_MS);
}

export function renderNames(scanned: boolean): void {
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

export async function scanNames(): Promise<string> {
  brokenList = await listBrokenNames();
  renderNames(true);
  return brokenList.length === 0 ? "No broken names" : brokenCount();
}

export async function deleteNames(): Promise<string> {
  const removed = await deleteBrokenNames();
  brokenList = [];
  renderNames(true);
  return `Deleted ${removed} broken ${removed === 1 ? "name" : "names"}`;
}

export async function insertTocSheet(): Promise<string> {
  await insertToc();
  await refreshSheets();
  return "Contents sheet updated";
}
