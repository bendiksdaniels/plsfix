// What an overlay painted over: the fill every covered cell had before, the map
// that outlives a single toggle, and the copy that rides along in a workbook
// setting so a reopened file puts the fills back before they can be mistaken
// for model formatting.
//
// Each overlay - the audit stripes and the linked-cell highlight - owns a store
// under its own setting key, so restoring one never writes over what the other
// remembers. Reading the fills stays with the caller: one overlay snapshots a
// single selection, the other every anchored range in one batch.

import { applyFillKey, fillKey, writeRuns } from "./internal";

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

// Nothing but a list of snapshots is worth restoring, and neither half of a
// corrupt setting may throw: the boot restore is the only thing that can take
// last session's stripes off, and its caller swallows what it throws.
function parseSnapshots(raw: string): FillSnapshot[] {
  try {
    const parsed: unknown = JSON.parse(raw);
    return Array.isArray(parsed) ? (parsed as FillSnapshot[]) : [];
  } catch {
    return [];
  }
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

export class FillStore {
  private readonly snapshots = new Map<string, FillSnapshot>();

  // The setting key is the store's identity: two overlays never share one. The
  // label is how the *other* overlay names this one when it refuses to paint.
  constructor(
    private readonly setting: string,
    private readonly label: string,
  ) {
    stores.push(this);
  }

  // Refuses while another overlay owns the fills, naming both the flow that
  // asked and the overlay standing in its way.
  requireSoleOwner(stage: string): void {
    const other = stores.find((store) => store !== this && store.painted);
    if (other) throw new Error(`${stage}: turn ${other.label} off first`);
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

  persist(context: Excel.RequestContext): void {
    const json = JSON.stringify([...this.snapshots.values()]);
    context.workbook.settings.add(
      this.setting,
      json.length > SETTING_MAX ? "" : json,
    );
  }

  // Puts every remembered fill back and forgets it. Nothing remembered is not
  // an empty write: the caller may be about to paint in the same batch.
  async restore(context: Excel.RequestContext): Promise<boolean> {
    if (this.snapshots.size === 0) return false;
    const snapshots = [...this.snapshots.values()];
    this.snapshots.clear();
    return this.writeBack(context, snapshots);
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
