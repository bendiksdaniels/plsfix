// What a PowerPoint pane button acts on: how a row is keyed, which rows a
// selection covers, and the lines an update reports. Pure maths over the state
// main.ts holds - no DOM, no Office.js and no relay, so every rule is testable
// without a host.

import { sourceLabel } from "../link/model";
import type { FoundLink } from "./host";
import type { LinkRow, UpdateSummary } from "./links";
import type { LinkRowView } from "./views";

// Slide plus shape: a copied link keeps its id and token, so identity in the
// list is the shape it lives on, never the link id.
export function rowKey(found: FoundLink): string {
  return `${found.slideId}/${found.shapeId}`;
}

export function toRowViews(
  rows: LinkRow[],
  selected: ReadonlySet<string>,
): LinkRowView[] {
  return rows.map((row) => {
    const key = rowKey(row.found);
    return {
      key,
      slide: row.found.slideIndex + 1,
      label: sourceLabel(row.found.tag.src, row.found.tag.kind),
      source: row.found.tag.src.workbook,
      status: row.status,
      pushedAt: row.pushedAt,
      selected: selected.has(key),
    };
  });
}

export function selectedRows(
  rows: LinkRow[],
  selected: ReadonlySet<string>,
): LinkRow[] {
  return rows.filter((row) => selected.has(rowKey(row.found)));
}

// "Update this slide" acts on the slide PowerPoint reports as active (see
// host.ts activeSlideId), never a tick: which rows sit on it is pure lookup.
export function slideRows(rows: LinkRow[], slideId: string): LinkRow[] {
  return rows.filter((row) => row.found.slideId === slideId);
}

// Only keys still in the deck survive a rescan: a shape someone deleted, or a
// link just broken, must not keep voting in the next "update selected".
export function pruneSelection(rows: LinkRow[], selected: Set<string>): void {
  const live = new Set(rows.map((row) => rowKey(row.found)));
  for (const key of selected) if (!live.has(key)) selected.delete(key);
}

export function requireSelection(rows: LinkRow[]): LinkRow[] {
  if (rows.length === 0) throw new Error("Tick a link in the list first.");
  return rows;
}

// The toast says how many did what; the details say which ones and why - a
// failure line per link, then every workbook a link now points at instead.
export function updateDetails(summary: UpdateSummary): string | undefined {
  const lines = [
    ...summary.failures,
    ...summary.sourceChanges.map((change) => `Source changed: ${change}`),
  ];
  return lines.length > 0 ? lines.join("\n") : undefined;
}
