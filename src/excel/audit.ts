// Formula audit overlay: stripes cells by precedent/dependent pattern (via the
// pure audit module) and remembers every fill it painted so it can restore the
// original formatting exactly. Snapshots persist in workbook settings, so a
// reload does not strand the paint.

import {
  applyFillKey,
  BASE_WHITE,
  fillKey,
  SELECTION_CELL_CAP,
  writeRuns,
} from "./internal";
import { parseAddress } from "./shared";
import { type AuditMark, auditGrid } from "../audit";
import { type CellValue } from "../model";
import { getActiveSettings, tint } from "../settings";

const LONE_FILL = "#E8B4B4";

interface FillSnapshot {
  sheetId: string;
  address: string;
  cells: string[][];
}

// The overlay owns nothing it did not paint: every fill it covers is stored here
// first and written back verbatim. It stays separate from SMT Undo because the
// overlay is a toggle the modeller turns off again, not an edit to the model.
const fillSnapshots = new Map<string, FillSnapshot>();

// The paint is saved with the file while this map dies with the runtime, so the
// snapshot rides along in workbook settings and startup restores it before the
// stripes can be mistaken for model formatting or re-snapshotted as original.
const OVERLAY_SETTING = "smtAuditOverlay";
const OVERLAY_SETTING_MAX = 400_000;

function persistOverlaySetting(context: Excel.RequestContext): void {
  const json = JSON.stringify([...fillSnapshots.values()]);
  context.workbook.settings.add(
    OVERLAY_SETTING,
    json.length > OVERLAY_SETTING_MAX ? "" : json,
  );
}

export async function restorePersistedOverlay(): Promise<boolean> {
  return Excel.run(async (context) => {
    const setting =
      context.workbook.settings.getItemOrNullObject(OVERLAY_SETTING);
    setting.load("isNullObject,value");
    await context.sync();
    if (setting.isNullObject || !setting.value) return false;

    let snapshots: FillSnapshot[] = [];
    try {
      snapshots = JSON.parse(String(setting.value)) as FillSnapshot[];
    } catch {
      snapshots = [];
    }

    const pending = snapshots.map((snapshot) => ({
      snapshot,
      sheet: context.workbook.worksheets.getItemOrNullObject(snapshot.sheetId),
    }));
    for (const { sheet } of pending) sheet.load("isNullObject");
    await context.sync();

    let restored = false;
    for (const { snapshot, sheet } of pending) {
      if (sheet.isNullObject) continue;
      writeRuns(sheet.getRange(snapshot.address), snapshot.cells, applyFillKey);
      restored = true;
    }
    context.workbook.settings.add(OVERLAY_SETTING, "");
    await context.sync();
    return restored;
  });
}

function overlayKey(mark: AuditMark, patternColor: string): string | null {
  switch (mark) {
    case "horizontal":
      return [Excel.FillPattern.lightHorizontal, BASE_WHITE, patternColor].join(
        "|",
      );
    case "vertical":
      return [Excel.FillPattern.lightVertical, BASE_WHITE, patternColor].join(
        "|",
      );
    case "both":
      return [Excel.FillPattern.crissCross, BASE_WHITE, patternColor].join("|");
    case "lone":
      return [Excel.FillPattern.solid, LONE_FILL, LONE_FILL].join("|");
    case "none":
      return null;
  }
}

export async function snapshotFills(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<void> {
  const properties = range.getCellProperties({
    format: { fill: { color: true, pattern: true, patternColor: true } },
  });
  const sheet = range.worksheet;
  sheet.load("id");
  range.load("address");
  await context.sync();

  const address = parseAddress(range.address).address;
  fillSnapshots.set(`${sheet.id}!${address}`, {
    sheetId: sheet.id,
    address,
    cells: properties.value.map((row) =>
      row.map((cell) => fillKey(cell.format?.fill)),
    ),
  });
}

export async function restoreFills(
  context: Excel.RequestContext,
): Promise<void> {
  if (fillSnapshots.size === 0) return;

  const pending = [...fillSnapshots.values()].map((snapshot) => ({
    snapshot,
    // Sheet id rather than name, so a rename between paint and restore is fine.
    sheet: context.workbook.worksheets.getItemOrNullObject(snapshot.sheetId),
  }));
  fillSnapshots.clear();

  for (const { sheet } of pending) sheet.load("isNullObject");
  await context.sync();

  for (const { snapshot, sheet } of pending) {
    if (sheet.isNullObject) continue;
    writeRuns(sheet.getRange(snapshot.address), snapshot.cells, applyFillKey);
  }
  context.workbook.settings.add(OVERLAY_SETTING, "");
  await context.sync();
}

export async function toggleAuditOverlay(): Promise<boolean> {
  return Excel.run(async (context) => {
    const selected = context.workbook.getSelectedRange();
    selected.load("rowCount,columnCount");
    await context.sync();

    // One cell says "check this block", not "check this cell".
    let target = selected;
    if (selected.rowCount === 1 && selected.columnCount === 1) {
      target = selected.getSurroundingRegion();
      target.load("rowCount,columnCount");
      await context.sync();
    }
    if (target.rowCount * target.columnCount > SELECTION_CELL_CAP) {
      throw new Error("The audit overlay supports up to 5,000 cells at once.");
    }

    const sheet = target.worksheet;
    sheet.load("id");
    target.load("address,formulasR1C1");
    await context.sync();

    const key = `${sheet.id}!${parseAddress(target.address).address}`;
    const wasOn = fillSnapshots.has(key);
    // Put old fills back before reading new ones, or the next snapshot would
    // capture our own paint over an overlapping range.
    await restoreFills(context);
    if (wasOn) return false;

    await snapshotFills(context, target);

    const patternColor = tint(getActiveSettings().primary, 0.55);
    const marks = auditGrid(target.formulasR1C1 as CellValue[][]);
    writeRuns(
      target,
      marks.map((row) => row.map((mark) => overlayKey(mark, patternColor))),
      applyFillKey,
    );
    persistOverlaySetting(context);
    await context.sync();
    return true;
  });
}
