// Every flow that turns the current selection into a new link: the range as a
// picture, the same range as a native table, or one cell as text. Owns the
// rules that refuse a selection - the cell cap, the table grid, the one-cell
// text rule - and they all run before anything is anchored, so a refused
// export leaves the workbook exactly as it was.

import { randomBytes } from "../link/crypto";
import {
  anchorName,
  newLinkId,
  overTableCap,
  sourceLabel,
  TABLE_TOO_BIG,
  TEXT_MAX_CHARS,
  TEXT_TOO_LONG,
} from "../link/model";
import type { RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import { SELECTION_CELL_CAP, selectedSingleRange } from "./internal";
import {
  createRangeAnchor,
  newEntry,
  readRegistry,
  requireImageApi,
  sourceOf,
  workbookName,
  type ResolvedSource,
} from "./link-anchors";
import { exclusive } from "./link-lock";
import { renderAnchored } from "./link-render";
import { publish, type NewLink } from "./link-record";
import { parseAddress } from "./shared";

export interface ExportResult {
  id: string;
  label: string;
  // What the toast adds after the label: for a chart Excel could not
  // describe, the sentence PowerPoint will repeat beside the picture.
  note?: string;
}

// What the selection is asked to become. A picture is what every host can
// draw, so it is the answer the other two point at when they refuse.
export type ExportKind = "range" | "table" | "text";

// Only a text link reads what the cell displays, so only it pays for that
// property on the batch that measures the selection.
const RANGE_PROPS = "address,cellCount,rowCount,columnCount,worksheet/name";

// The selected range as a linked picture.
export async function exportSelection(
  ws: Workspace,
  relay: RelayApi,
): Promise<ExportResult> {
  return exportRange(ws, relay, "range");
}

// The same range as an editable PowerPoint table: same anchor, same registry
// entry, same inbox note - only the render differs.
export async function exportSelectionAsTable(
  ws: Workspace,
  relay: RelayApi,
): Promise<ExportResult> {
  return exportRange(ws, relay, "table");
}

// One cell's displayed text as a text box on the slide: same anchor, same
// registry entry, same inbox note as a picture of it.
export async function exportSelectionAsText(
  ws: Workspace,
  relay: RelayApi,
): Promise<ExportResult> {
  return exportRange(ws, relay, "text");
}

// Every cap is checked before anything is anchored, so a selection too big to
// send leaves the workbook exactly as it was.
function requireExportable(range: Excel.Range, kind: ExportKind): void {
  if (range.cellCount > SELECTION_CELL_CAP) {
    throw new Error(
      `Export supports up to ${SELECTION_CELL_CAP.toLocaleString()} selected cells at once.`,
    );
  }
  if (kind === "table" && overTableCap(range.rowCount, range.columnCount)) {
    throw new Error(TABLE_TOO_BIG);
  }
  if (kind === "text") requireTextCell(range);
}

// One cell only: a merged area arrives as a single range of several cells, and
// a picture of it says the same thing without guessing which cell wins.
function requireTextCell(range: Excel.Range): void {
  if (range.cellCount !== 1) {
    throw new Error(
      "Select one cell for a text link (merged cells: export as a picture).",
    );
  }
  const text = range.text[0]?.[0] ?? "";
  if (text.trim() === "") throw new Error("The cell is empty.");
  if (text.length > TEXT_MAX_CHARS) throw new Error(TEXT_TOO_LONG);
}

// Every flow that rewrites the registry runs through the shared link queue: the
// read, the upload and the write-back are one critical section, or a push that
// began earlier puts its own copy of the registry back over this new link.
async function exportRange(
  ws: Workspace,
  relay: RelayApi,
  kind: ExportKind,
): Promise<ExportResult> {
  requireImageApi();
  const workbook = await workbookName();
  return exclusive(kind === "range" ? "export" : `export ${kind}`, () =>
    Excel.run(async (context) => {
      const registry = await readRegistry(context);
      const range = await selectedSingleRange(context, "export");
      range.load(kind === "text" ? `${RANGE_PROPS},text` : RANGE_PROPS);
      await context.sync();
      requireExportable(range, kind);

      const resolved: ResolvedSource = {
        kind,
        sheet: range.worksheet.name,
        ref: parseAddress(range.address).address,
        range,
      };
      const id = newLinkId(randomBytes);
      const anchor = anchorName(id);
      const src = sourceOf(workbook, anchor, resolved);
      const entry = newEntry(id, kind, anchor, sourceLabel(src, kind));
      // Anchor and render in one batch, before any network call: a selection
      // that changes during the upload cannot make the two describe different
      // objects. The render owns the anchor from here on, so a picture that
      // never arrives takes the name with it.
      const named = createRangeAnchor(context, range, anchor);
      const release = () => named.delete();
      const render = await renderAnchored(
        context,
        resolved,
        entry.label,
        release,
      );

      const link: NewLink = { entry, src, render, registry, release };
      await publish(context, link, ws, relay);
      return { id, label: entry.label };
    }),
  );
}
