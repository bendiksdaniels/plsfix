// Workbook tools: the sheet explorer (go to, show/hide), the six sheet tools
// above it, the broken-name scrubber and the contents-sheet writer. The "click
// again to confirm" arm on the delete button lapses on its own, owned by
// src/ui/confirm.ts. Office.js only reaches here through ../excel.
//
// Every sheet tool re-renders the explorer, and every one of them says in its
// own line that it changed the sheet list: those changes are outside pls,fix
// Undo (see the header of src/excel/workbook.ts).

import {
  activateSheet,
  burySheet,
  deleteBrokenNames,
  insertToc,
  listBrokenNames,
  listSheets,
  moveSheet,
  setSheetsVisibility,
  setSheetVisibility,
  type SheetEntry,
  type SheetMove,
  showOnlySheet,
} from "../excel";
import { armConfirm, type ConfirmButton } from "../ui/confirm";
import { getElement } from "../ui/dom";
import { errorMessage, guard, isExcelReady, toast } from "./shared";

let brokenList: string[] = [];
let nameDeleteConfirm: ConfirmButton | undefined;

// Lazy for the same reason src/pane/styles-panel.ts's own confirm is: built
// on first use, run unused since main.ts still owns the deleteNames() call
// (this module cannot import dispatch.ts without a cycle).
function deleteNamesConfirm(): ConfirmButton {
  nameDeleteConfirm ??= armConfirm(
    getElement<HTMLButtonElement>("delete-names"),
    () => undefined,
    { label: deleteLabel },
  );
  return nameDeleteConfirm;
}

export function isDeleteArmed(): boolean {
  return deleteNamesConfirm().isArmed();
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
  deleteNamesConfirm().disarm();
}

// Deleting a name cannot be undone by us or by Excel, so the first click only
// arms the button and the arming lapses on its own.
export function armDelete(): void {
  deleteNamesConfirm().arm();
}

// Nothing else ever disables this button - only setBusy's blanket pass does,
// so put back after it - whether the scrubber has been run this session or
// not changes what it says, never whether it is disabled.
export function syncDeleteNamesButton(): void {
  getElement<HTMLButtonElement>("delete-names").disabled = false;
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

// The sheet tools. Each one acts, re-renders the explorer and hands back the
// line the toast shows; the counts are plural-correct because a modeller reads
// them as a receipt for a change Excel's own undo cannot take back.
function sheetWord(count: number): string {
  return count === 1 ? "sheet" : "sheets";
}

function includeVeryHidden(): boolean {
  return getElement<HTMLInputElement>("sheets-very-hidden").checked;
}

// The buried sheets a pass left behind. "1 sheet shown" on its own reads as
// "that was all of them" on a workbook still hiding two the tick would find.
function buriedNote(buried: number): string {
  if (buried === 0) return "";
  const many = buried > 1;
  const count = `${String(buried)} more ${many ? "are" : "is"} very hidden`;
  return `. ${count}: tick "include very hidden" for ${many ? "them" : "it"}`;
}

export async function unhideAllSheets(): Promise<string> {
  const { shown, buried } = await setSheetsVisibility(includeVeryHidden());
  await refreshSheets();
  if (shown > 0) {
    return `${String(shown)} ${sheetWord(shown)} shown${buriedNote(buried)}`;
  }
  // "No hidden sheets" would be a lie while the tick would still find some.
  if (buried > 0) {
    return 'No hidden sheets to show. Tick "include very hidden" for the buried ones.';
  }
  return "No hidden sheets to show";
}

export async function showOnlyThisSheet(): Promise<string> {
  const { name, hidden } = await showOnlySheet();
  await refreshSheets();
  if (hidden === 0) return `${name} was already the only visible sheet`;
  return `Only ${name} is visible now: ${String(hidden)} ${sheetWord(hidden)} hidden`;
}

export async function buryThisSheet(): Promise<string> {
  const name = await burySheet();
  await refreshSheets();
  return `${name} is very hidden now. Unhide all brings it back with the tick on`;
}

export async function moveThisSheet(direction: SheetMove): Promise<string> {
  const { name, position } = await moveSheet(direction);
  await refreshSheets();
  return `${name} is now sheet ${String(position + 1)}`;
}
