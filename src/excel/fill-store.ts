// What an overlay painted over: the fill every covered cell had before, the map
// that outlives a single toggle, and the copy that rides along in a workbook
// setting so a reopened file puts the fills back before they can be mistaken
// for model formatting. Also the one rule a fill read back from Excel is
// understood by (readFill), and the key a remembered fill travels as.
//
// Each overlay - the audit stripes and the linked-cell highlight - owns a store
// under its own setting key, so restoring one never writes over what the other
// remembers. Reading the fills stays with the caller: one overlay snapshots a
// single selection, the other every anchored range in one batch.

import { parseRect } from "../link/geometry";
import { BASE_WHITE, onTheWeb, writeRuns } from "./internal";

export interface FillSnapshot {
  sheetId: string;
  // Sheet-local, the way worksheet.getRange wants it.
  address: string;
  cells: string[][];
}

// A setting is saved inside the workbook, so a snapshot past this size is
// dropped rather than bloating every save: this session's toggle is then the
// only way back to the original fills.
const SETTING_MAX = 400_000;

function snapshotKey(sheetId: string, address: string): string {
  return `${sheetId}!${address}`;
}

// A list entry is a snapshot only with all three parts in place: a bare number
// or a half-written object would reach worksheet.getRange with no address.
function isSnapshot(entry: unknown): entry is FillSnapshot {
  if (typeof entry !== "object" || entry === null) return false;
  const snapshot = entry as Partial<FillSnapshot>;
  return (
    typeof snapshot.sheetId === "string" &&
    typeof snapshot.address === "string" &&
    Array.isArray(snapshot.cells)
  );
}

// Nothing but a list of snapshots is worth restoring, and neither half of a
// corrupt setting may throw: the boot restore is the only thing that can take
// last session's stripes off, and its caller swallows what it throws.
function parseSnapshots(raw: string): FillSnapshot[] {
  try {
    const parsed: unknown = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter(isSnapshot) : [];
  } catch {
    return [];
  }
}

const NO_FILL = "none";

/** A fill as Excel can take it back: its pattern and the colours it named. */
export interface ReadFill {
  pattern: string;
  color?: string;
  patternColor?: string;
}

/** A fill as either API reads it: getCellProperties or format.fill. */
export interface FillShape {
  color?: string | null;
  pattern?: string | null;
  patternColor?: string | null;
}

/**
 * The one way a fill read back from Excel is understood; null is no fill.
 * Excel for the web answers a solid fill with a null pattern and an unfilled
 * cell with colour "" (rig 27.09); Excel for Mac answers its unfilled cell
 * with a null pattern over white (lessons.md). So a null pattern is a solid
 * fill of its colour on the web alone, and no fill on every other host.
 */
export function readFill(
  fill: FillShape | undefined,
  web = onTheWeb(),
): ReadFill | null {
  const pattern = fill?.pattern ?? null;
  if (pattern === "None") return null;
  if (pattern === null) {
    // The web never names a solid fill's pattern colour at all (rig 27.09),
    // but a solid fill has only one visible colour: naming it here too, the
    // way fillKey's own "unnamed" default already treats a solid fill, keeps
    // every caller - not only fillKey - from reading this fill as one with no
    // pattern colour at all. undo.ts's settableFill is the one that mattered:
    // it only writes a field readFill actually named, so a captured solid
    // fill with no named pattern colour left whatever was under it (a
    // pinstripe band's tint, say) in place after a restore.
    return web && fill?.color
      ? { pattern: "Solid", color: fill.color, patternColor: fill.color }
      : null;
  }
  const read: ReadFill = { pattern };
  if (fill?.color) read.color = fill.color;
  if (fill?.patternColor) read.patternColor = fill.patternColor;
  return read;
}

// Colour, pattern and pattern colour together, so a modeller's own striped fill
// comes back exactly as it was. A solid fill whose pattern colour Excel never
// named (the web never does) keys it as its own colour: the shape every
// overlay and the pinstripes paint with, so a band still reads as a band.
export function fillKey(fill: FillShape | undefined): string {
  const read = readFill(fill);
  if (!read) return NO_FILL;
  const color = read.color ?? BASE_WHITE;
  const unnamed = read.pattern === "Solid" ? color : BASE_WHITE;
  return [read.pattern, color, read.patternColor ?? unnamed].join("|");
}

export function applyFillKey(block: Excel.Range, key: string): void {
  const { fill } = block.format;
  if (key === NO_FILL) {
    fill.clear();
    return;
  }
  const [pattern, color, patternColor] = key.split("|");
  // Colour first: setting it on an unfilled cell would otherwise force Solid.
  fill.color = color ?? BASE_WHITE;
  fill.pattern = pattern as Excel.FillPattern;
  fill.patternColor = patternColor ?? BASE_WHITE;
}

// The colour half of a fill key, the same one format.fill.color would have
// read before an overlay ever painted: null for "none" (no fill at all), the
// colour otherwise, pattern and pattern colour dropped since a table cell
// only ever wants one solid colour.
export function fillKeyColor(key: string): string | null {
  if (key === NO_FILL) return null;
  return key.split("|")[1] ?? null;
}

// Colour, pattern and pattern colour together, so a modeller's own striped fill
// comes back exactly as it was.
export function requestFills(
  range: Excel.Range,
): OfficeExtension.ClientResult<Excel.CellProperties[][]> {
  return range.getCellProperties({
    format: { fill: { color: true, pattern: true, patternColor: true } },
  });
}

// Call after the sync the request was queued in.
export function fillGrid(
  properties: OfficeExtension.ClientResult<Excel.CellProperties[][]>,
): string[][] {
  return properties.value.map((row) =>
    row.map((cell) => fillKey(cell.format?.fill)),
  );
}

// Every store there is. Two overlays must never both hold fills: the second to
// paint would snapshot the first's colour as the modeller's own formatting, and
// on the way out one of them would hand that colour back as if it belonged to
// the model. Whoever paints asks here first, so the refusal is symmetric by
// construction rather than by one of them remembering to check the other.
const stores: FillStore[] = [];

/**
 * Refuses while an overlay owns the fills, naming the flow that asked and the
 * overlay standing in its way. `except` is the store the caller owns itself, so
 * the two overlays ask only about each other; a flow that paints without a
 * store of its own - the pinstripes - passes none and both hold it off.
 */
export function requireNoOverlayOwner(stage: string, except?: FillStore): void {
  const owner = stores.find((store) => store !== except && store.painted);
  if (owner) throw new Error(`${stage}: turn ${owner.label} off first`);
}

/**
 * The colour a painted store remembers for one cell, across every overlay:
 * undefined when neither owns it, which tells the caller its own live read
 * already is the true fill. Checked before a table export trusts a cell's
 * live fill, so an overlay's tint - never the model's own formatting - is
 * never the one that ships. Pure lookup against in-memory state: no sync.
 */
export function originalFillColor(
  sheetId: string,
  row: number,
  column: number,
): string | null | undefined {
  for (const store of stores) {
    if (!store.painted) continue;
    const color = store.originalColorAt(sheetId, row, column);
    if (color !== undefined) return color;
  }
  return undefined;
}

export class FillStore {
  private readonly snapshots = new Map<string, FillSnapshot>();

  // The setting key is the store's identity: two overlays never share one. The
  // label is how anything else names this one when it refuses to paint.
  constructor(
    private readonly setting: string,
    readonly label: string,
  ) {
    stores.push(this);
  }

  // Refuses while another overlay owns the fills, naming both the flow that
  // asked and the overlay standing in its way.
  requireSoleOwner(stage: string): void {
    requireNoOverlayOwner(stage, this);
  }

  // Whether this overlay is on: it owns a fill exactly while it remembers what
  // was under it.
  get painted(): boolean {
    return this.snapshots.size > 0;
  }

  has(sheetId: string, address: string): boolean {
    return this.snapshots.has(snapshotKey(sheetId, address));
  }

  remember(sheetId: string, address: string, cells: string[][]): void {
    this.snapshots.set(snapshotKey(sheetId, address), {
      sheetId,
      address,
      cells,
    });
  }

  // The colour this store remembers under one cell (1-based row and column,
  // A1-style): null when it owned the cell but the original there was no
  // fill, undefined when none of this store's snapshots cover it at all.
  originalColorAt(
    sheetId: string,
    row: number,
    column: number,
  ): string | null | undefined {
    for (const snapshot of this.snapshots.values()) {
      if (snapshot.sheetId !== sheetId) continue;
      const rect = parseRect(snapshot.address);
      if (!rect) continue;
      const { top, left, bottom, right } = rect;
      if (row < top || row > bottom || column < left || column > right) {
        continue;
      }
      const key = snapshot.cells[row - top]?.[column - left];
      return key === undefined ? undefined : fillKeyColor(key);
    }
    return undefined;
  }

  persist(context: Excel.RequestContext): void {
    const json = JSON.stringify([...this.snapshots.values()]);
    context.workbook.settings.add(
      this.setting,
      json.length > SETTING_MAX ? "" : json,
    );
  }

  // Puts every remembered fill back and forgets it. Nothing remembered is not
  // an empty write: the caller may be about to paint in the same batch.
  // Forgets only once the write-back actually lands, or once it is certain
  // it never touched the sheet at all: a snapshot cleared ahead of a sync
  // that may have already applied its writes (a dropped connection, say) is
  // the original fill lost for good, with the tint still on the sheet and
  // nothing left that remembers what was under it. A sheet that refuses the
  // write outright (AccessDenied - protection.ts's own distinguishing
  // signal) never took it in the first place, so there is nothing under the
  // tint to lose by forgetting: keeping the record there instead would
  // leave the store believing it owns a cell forever, until the sheet is
  // unprotected.
  async restore(context: Excel.RequestContext): Promise<boolean> {
    if (this.snapshots.size === 0) return false;
    const snapshots = [...this.snapshots.values()];
    try {
      const restored = await this.writeBack(context, snapshots);
      this.snapshots.clear();
      return restored;
    } catch (error) {
      if ((error as { code?: string }).code === Excel.ErrorCodes.accessDenied) {
        this.snapshots.clear();
      }
      throw error;
    }
  }

  // Startup: the map died with the pane while the paint was saved with the
  // file, so the snapshot in the setting is what puts the fills back.
  async restorePersisted(context: Excel.RequestContext): Promise<boolean> {
    const setting = context.workbook.settings.getItemOrNullObject(this.setting);
    setting.load("isNullObject,value");
    await context.sync();
    if (setting.isNullObject || !setting.value) return false;
    return this.writeBack(context, parseSnapshots(String(setting.value)));
  }

  // Sheets are found by id rather than name, so a rename between paint and
  // restore is fine; a sheet that is gone is skipped, not an error.
  private async writeBack(
    context: Excel.RequestContext,
    snapshots: FillSnapshot[],
  ): Promise<boolean> {
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
    context.workbook.settings.add(this.setting, "");
    await context.sync();
    return restored;
  }
}
