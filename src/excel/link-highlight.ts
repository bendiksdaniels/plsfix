// "Highlight linked cells": a subtle brand-tinted fill over every anchored
// range in this workbook, so a modeller sees at a glance which blocks feed
// PowerPoint, and the fills exactly as they were when the toggle goes off. The
// snapshot rides along in a workbook setting, so a file reopened with the paint
// still on gets its own formatting back before the tint is mistaken for it.
//
// Order rule: the highlight and the audit overlay both own the fills they paint
// and both remember what was under them, so only one may be on at a time. The
// highlight refuses to paint while the overlay is on; turn the highlight off
// before toggling the audit overlay, or the overlay would snapshot this tint as
// the modeller's own formatting and hand it back as such.
//
// Charts are never highlighted: a chart anchor is the chart's own name, not a
// range, so there are no cells under it to tint.

import { auditOverlayOn } from "./audit";
import { FillStore, fillGrid, requestFills } from "./fill-store";
import { applyFillKey, SELECTION_CELL_CAP } from "./internal";
import { readRegistry, resolveSource } from "./link-anchors";
import { parseAddress } from "./shared";
import { activeTheme, tint } from "../settings";

export const HIGHLIGHT_SETTING = "SMT_LINK_HIGHLIGHT";

// Light enough to read a model through, dark enough to find on a white grid.
const HIGHLIGHT_TINT = 0.85;

const highlight = new FillStore(HIGHLIGHT_SETTING);

interface Anchored {
  range: Excel.Range;
  sheetId: string;
  address: string;
  cells: number;
}

// The link colour of the active palette, tinted: the fill says "this is linked"
// in the same colour the pane uses for links everywhere else.
function highlightKey(): string {
  const color = tint(activeTheme().linkFont, HIGHLIGHT_TINT);
  return [Excel.FillPattern.solid, color, color].join("|");
}

// Range links only. A chart is skipped by kind, and a range whose rows were
// deleted resolves to nothing: the Links tab reports that one as "Source
// missing" and owns its removal, so the highlight leaves it alone.
async function anchoredRanges(
  context: Excel.RequestContext,
): Promise<Excel.Range[]> {
  const registry = await readRegistry(context);
  const ranges: Excel.Range[] = [];
  for (const entry of registry.links) {
    if (entry.kind !== "range") continue;
    const resolved = await resolveSource(context, entry);
    if (resolved === null || resolved.kind !== "range") continue;
    ranges.push(resolved.range);
  }
  return ranges;
}

// Sheet id rather than name, so a rename between paint and restore is fine, and
// the cell counts in one batch, so the cap is decided before anything is read
// cell by cell.
async function measure(
  context: Excel.RequestContext,
  ranges: Excel.Range[],
): Promise<Anchored[]> {
  for (const range of ranges) {
    range.load("address,cellCount,worksheet/id");
  }
  await context.sync();

  return ranges.map((range) => ({
    range,
    sheetId: range.worksheet.id,
    address: parseAddress(range.address).address,
    cells: range.cellCount,
  }));
}

// Every fill is read before any is painted: two links may overlap, and a
// snapshot taken after a neighbour was tinted would hand back our own paint.
async function paint(
  context: Excel.RequestContext,
  anchored: Anchored[],
): Promise<void> {
  const properties = anchored.map((one) => requestFills(one.range));
  await context.sync();

  anchored.forEach((one, index) => {
    const grid = properties[index];
    if (grid) highlight.remember(one.sheetId, one.address, fillGrid(grid));
  });

  const key = highlightKey();
  for (const one of anchored) applyFillKey(one.range, key);
  highlight.persist(context);
  await context.sync();
}

// Returns the state the toggle left behind: true when the workbook is painted.
export async function toggleLinkHighlight(): Promise<boolean> {
  if (auditOverlayOn()) {
    throw new Error("highlight: turn the audit overlay off first");
  }
  return Excel.run(async (context) => {
    if (await highlight.restore(context)) return false;

    const anchored = await measure(context, await anchoredRanges(context));
    if (anchored.length === 0) {
      throw new Error("highlight: no linked ranges in this workbook");
    }
    const cells = anchored.reduce((total, one) => total + one.cells, 0);
    if (cells > SELECTION_CELL_CAP) {
      throw new Error(
        `highlight: linked ranges exceed ${SELECTION_CELL_CAP.toLocaleString()} cells`,
      );
    }

    await paint(context, anchored);
    return true;
  });
}

// Boot: the map died with the pane while the paint was saved with the file, so
// last session's fills go back before they can be mistaken for model
// formatting. Returns whether anything was put back; the toggle reads off.
export async function restoreLinkHighlight(): Promise<boolean> {
  return Excel.run((context) => highlight.restorePersisted(context));
}
