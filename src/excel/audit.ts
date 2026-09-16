// Formula audit overlay: stripes cells by precedent/dependent pattern (via the
// pure audit module) and remembers every fill it painted so it can restore the
// original formatting exactly. The snapshot store persists in workbook
// settings, so a reload does not strand the paint.

import { FillStore, fillGrid, requestFills } from "./fill-store";
import {
  applyFillKey,
  BASE_WHITE,
  selectedSingleRange,
  SELECTION_CELL_CAP,
  writeRuns,
} from "./internal";
import { protectedNote, sheetProtected } from "./protection";
import { parseAddress } from "./shared";
import { type AuditMark, auditGrid } from "../audit";
import { type CellValue } from "../model";
import { getActiveSettings, tint } from "../settings";

const LONE_FILL = "#E8B4B4";

// One cap, one sentence, one place: every audit tool reads a grid under
// SELECTION_CELL_CAP, and each names itself so a refusal never blames a button
// the modeller did not press.
export function scanCapSentence(tool: string): string {
  return `${tool} supports up to ${SELECTION_CELL_CAP.toLocaleString()} cells at once.`;
}

// The overlay owns nothing it did not paint: every fill it covers is stored
// first and written back verbatim. It stays separate from pls,fix Undo because the
// overlay is a toggle the modeller turns off again, not an edit to the model.
// The paint is saved with the file while the map dies with the runtime, so the
// snapshot rides along in this workbook setting and startup restores it before
// the stripes can be mistaken for model formatting or re-snapshotted.
const OVERLAY_SETTING = "smtAuditOverlay";
const overlay = new FillStore(OVERLAY_SETTING, "the audit overlay");

// Read by the linked-cell highlight before it paints: two overlays cannot own
// the same fill, and a highlight painted over these stripes would be snapshot
// as the modeller's own formatting.
export function auditOverlayOn(): boolean {
  return overlay.painted;
}

export async function restorePersistedOverlay(): Promise<boolean> {
  return Excel.run((context) => overlay.restorePersisted(context));
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
    // A typed number reads as one obvious tinted cell rather than a stripe:
    // there is no second formula to weave a pattern against, only the tint
    // that says "this looks like it belongs to the formulas around it".
    case "typed":
      return [Excel.FillPattern.solid, patternColor, patternColor].join("|");
    case "none":
      return null;
  }
}

export async function snapshotFills(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<void> {
  const properties = requestFills(range);
  const sheet = range.worksheet;
  sheet.load("id");
  range.load("address");
  await context.sync();

  overlay.remember(
    sheet.id,
    parseAddress(range.address).address,
    fillGrid(properties),
  );
}

export async function restoreFills(
  context: Excel.RequestContext,
): Promise<void> {
  await overlay.restore(context);
}

const OVERLAY_STAGE = "Audit overlay";

let overlayNote: string | null = null;

// Read-once, the way lastUndoSkipped is: what the last toggle could not do.
// The pane shows it instead of the plain on/off line when it is not null.
export function lastAuditNote(): string | null {
  const note = overlayNote;
  overlayNote = null;
  return note;
}

export async function toggleAuditOverlay(): Promise<boolean> {
  // The linked-cell highlight owns fills of its own, and both stores hand back
  // what they covered: painting over the other one would give the modeller our
  // tint back as if it were their formatting.
  overlay.requireSoleOwner("audit overlay");
  overlayNote = null;
  return Excel.run(async (context) => {
    const selected = await selectedSingleRange(context, OVERLAY_STAGE);
    selected.load("rowCount,columnCount");
    await context.sync();

    // A protected sheet refuses every fill the overlay would write, and half a
    // painted overlay is worse than none: it is skipped and said so.
    if (await sheetProtected(context, selected.worksheet)) {
      overlayNote = protectedNote(OVERLAY_STAGE);
      return false;
    }

    // One cell says "check this block", not "check this cell".
    let target = selected;
    if (selected.rowCount === 1 && selected.columnCount === 1) {
      target = selected.getSurroundingRegion();
      target.load("rowCount,columnCount");
      await context.sync();
    }
    if (target.rowCount * target.columnCount > SELECTION_CELL_CAP) {
      throw new Error(scanCapSentence("The audit overlay"));
    }

    const sheet = target.worksheet;
    sheet.load("id");
    target.load("address,formulasR1C1");
    await context.sync();

    const wasOn = overlay.has(sheet.id, parseAddress(target.address).address);
    // Put old fills back before reading new ones, or the next snapshot would
    // capture our own paint over an overlapping range.
    await restoreFills(context);
    if (wasOn) return false;

    const patternColor = tint(getActiveSettings().primary, 0.55);
    const keys = auditGrid(target.formulasR1C1 as CellValue[][]).map((row) =>
      row.map((mark) => overlayKey(mark, patternColor)),
    );
    // Nothing to stripe is nothing to own: taking the fills anyway would leave
    // the overlay reading "on" over a block with no stripes on it, save a
    // snapshot in the file for nothing, and lock the linked-cell highlight out
    // until the modeller toggled an overlay they cannot see.
    if (keys.every((row) => row.every((key) => key === null))) {
      overlayNote = "no formula here to stripe";
      return false;
    }
    await snapshotFills(context, target);

    writeRuns(target, keys, applyFillKey);
    overlay.persist(context);
    await context.sync();
    return true;
  });
}
