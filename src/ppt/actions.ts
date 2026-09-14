// What a PowerPoint pane button acts on: how a row is keyed, which rows a
// selection covers, and the lines an update reports. Pure maths over the state
// main.ts holds - no DOM, no Office.js and no relay, so every rule is testable
// without a host.

import { sourceLabel } from "../link/model";
import { projectLabel } from "../link/project";
import type { FoundLink } from "./host";
import type { LinkRow, UpdateSummary } from "./links";
import type { LinkRowView } from "./views";

export type LinkFilterStatus = "all" | LinkRow["status"];

export interface LinkFilter {
  query: string;
  status: LinkFilterStatus;
  source?: string;
  slide?: number | "all";
  project?: string;
}

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
      project: row.found.tag.project,
      kind: row.found.tag.kind,
      status: row.status,
      pushedAt: row.pushedAt,
      selected: selected.has(key),
    };
  });
}

// Keep filtering in the state layer, rather than in the renderer, so a
// selected link stays selected when it is temporarily outside the current
// view. Search covers the labels people can see: source workbook, sheet/range,
// kind, status and the one-based slide number.
export function filterRows(rows: LinkRow[], filter: LinkFilter): LinkRow[] {
  const query = filter.query.trim().toLocaleLowerCase();
  return rows.filter((row) => {
    if (filter.status !== "all" && row.status !== filter.status) return false;
    if (
      filter.source !== undefined &&
      filter.source !== "all" &&
      row.found.tag.src.workbook !== filter.source
    ) {
      return false;
    }
    if (
      filter.project !== undefined &&
      filter.project !== "all" &&
      (row.found.tag.project ?? "") !== filter.project
    ) {
      return false;
    }
    if (
      filter.slide !== undefined &&
      filter.slide !== "all" &&
      row.found.slideIndex + 1 !== filter.slide
    ) {
      return false;
    }
    if (query === "") return true;
    const values = [
      sourceLabel(row.found.tag.src, row.found.tag.kind),
      row.found.tag.src.workbook,
      row.found.tag.project ?? "",
      row.found.tag.kind,
      row.status,
      String(row.found.slideIndex + 1),
    ];
    return values.some((value) => value.toLocaleLowerCase().includes(query));
  });
}

export function linkSources(rows: LinkRow[]): string[] {
  return [...new Set(rows.map((row) => row.found.tag.src.workbook))].sort();
}

export function linkProjects(rows: LinkRow[]): string[] {
  return [
    ...new Set(
      rows
        .map((row) => row.found.tag.project)
        .filter((name): name is string => name !== undefined),
    ),
  ].sort();
}

export function linkSlides(rows: LinkRow[]): number[] {
  return [...new Set(rows.map((row) => row.found.slideIndex + 1))].sort(
    (a, b) => a - b,
  );
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

// The message is the button's, not this function's: "Tick a link in the list
// first." reads wrong under a button that acts on several rows at once.
export function requireSelection(
  rows: LinkRow[],
  message = "Tick a link in the list first.",
): LinkRow[] {
  if (rows.length === 0) throw new Error(message);
  return rows;
}

// The toast says how many did what; the details say which ones and why - a
// failure line per link, then every workbook a link now points at instead,
// then every chart the deck holds as a picture and the reason.
export function scopedSummary(text: string, project?: string): string {
  if (project === undefined || project === "all") return text;
  const name = projectLabel(project === "" ? undefined : project);
  return `${text} (${name})`;
}

export function updateDetails(summary: UpdateSummary): string | undefined {
  const lines = [
    ...summary.failures,
    ...summary.sourceChanges.map((change) => `Source changed: ${change}`),
    ...summary.notes,
  ];
  return lines.length > 0 ? lines.join("\n") : undefined;
}
