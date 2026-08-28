// The Excel side of tracked links, one function per pane action: export a
// selection or the active chart, push every anchored source again, list what
// this workbook owns, jump back to a source and remove a link. The anchor - not
// the address it was created at - is what every later flow resolves through.

import { randomBytes } from "../link/crypto";
import {
  anchorName,
  newLinkId,
  sourceLabel,
  type Registry,
  type RegistryEntry,
} from "../link/model";
import type { RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import { hostSupports, SELECTION_CELL_CAP } from "./internal";
import {
  createChartAnchor,
  createRangeAnchor,
  forget,
  newEntry,
  publish,
  pushPayload,
  readRegistry,
  releaseAnchor,
  renderSource,
  resolveSource,
  sourceOf,
  workbookName,
  writeRegistry,
  type NewLink,
  type ResolvedSource,
} from "./link-anchors";
import { parseAddress } from "./shared";

export { workbookName };

export interface WorkbookLinkRow {
  entry: RegistryEntry;
  source: "ok" | "missing";
}

export interface ExportResult {
  id: string;
  label: string;
}

export interface PushSummary {
  pushed: number;
  missing: number;
  failed: number;
}

function entryOf(registry: Registry, id: string, stage: string): RegistryEntry {
  const entry = registry.links.find((link) => link.id === id);
  if (!entry) throw new Error(`${stage} ${id}: not in this workbook`);
  return entry;
}

export async function exportSelection(
  ws: Workspace,
  relay: RelayApi,
): Promise<ExportResult> {
  const workbook = await workbookName();
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address,cellCount,worksheet/name");
    await context.sync();
    if (range.cellCount > SELECTION_CELL_CAP) {
      throw new Error(
        `Export supports up to ${SELECTION_CELL_CAP.toLocaleString()} selected cells at once.`,
      );
    }

    const resolved: ResolvedSource = {
      kind: "range",
      sheet: range.worksheet.name,
      ref: parseAddress(range.address).address,
      range,
    };
    const png = await renderSource(context, resolved);
    const id = newLinkId(randomBytes);
    const anchor = anchorName(id);
    const src = sourceOf(workbook, anchor, resolved);
    const entry = newEntry(id, "range", anchor, sourceLabel(src, "range"));
    const link: NewLink = {
      entry,
      src,
      png,
      anchor: () => createRangeAnchor(context, range, anchor),
    };
    await publish(context, link, ws, relay);
    return { id, label: entry.label };
  });
}

// The same guard formatSelectedChart uses: the hosted office.js always defines
// the method, so the host's API set is what decides.
async function activeChart(
  context: Excel.RequestContext,
): Promise<Excel.Chart> {
  const callable = (context.workbook as unknown as Record<string, unknown>)
    .getActiveChartOrNullObject;
  if (typeof callable !== "function" || !hostSupports("1.9")) {
    throw new Error("Exporting a chart needs a newer Excel build.");
  }
  const chart = context.workbook.getActiveChartOrNullObject();
  chart.load("isNullObject");
  await context.sync();
  if (chart.isNullObject) throw new Error("Select a chart first.");

  chart.load("name,width,height,worksheet/name");
  await context.sync();
  return chart;
}

export async function exportActiveChart(
  ws: Workspace,
  relay: RelayApi,
): Promise<ExportResult> {
  const workbook = await workbookName();
  return Excel.run(async (context) => {
    const chart = await activeChart(context);
    const resolved: ResolvedSource = {
      kind: "chart",
      sheet: chart.worksheet.name,
      ref: chart.name,
      chart,
      width: chart.width,
      height: chart.height,
    };
    const png = await renderSource(context, resolved);
    const id = newLinkId(randomBytes);
    const anchor = anchorName(id);
    const src = sourceOf(workbook, anchor, resolved);
    const entry = newEntry(id, "chart", anchor, sourceLabel(src, "chart"));
    const link: NewLink = {
      entry,
      src,
      png,
      anchor: () => createChartAnchor(chart, anchor),
    };
    await publish(context, link, ws, relay);
    return { id, label: entry.label };
  });
}

export async function listWorkbookLinks(): Promise<WorkbookLinkRow[]> {
  return Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const rows: WorkbookLinkRow[] = [];
    for (const entry of registry.links) {
      const resolved = await resolveSource(context, entry);
      rows.push({ entry, source: resolved === null ? "missing" : "ok" });
    }
    return rows;
  });
}

// A push is a report, not an assertion: a source that is gone and a relay that
// refused are counted rather than thrown, so one bad link cannot stop the rest.
export async function pushLinks(
  ids: string[] | "all",
  relay: RelayApi,
): Promise<PushSummary> {
  const workbook = await workbookName();
  return Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const wanted = ids === "all" ? null : new Set(ids);
    if (wanted) for (const id of wanted) entryOf(registry, id, "push");

    const summary: PushSummary = { pushed: 0, missing: 0, failed: 0 };
    for (const entry of registry.links) {
      if (wanted && !wanted.has(entry.id)) continue;
      const resolved = await resolveSource(context, entry);
      if (resolved === null) {
        summary.missing += 1;
        continue;
      }
      try {
        const png = await renderSource(context, resolved);
        const src = sourceOf(workbook, entry.anchor, resolved);
        entry.rev = await pushPayload(entry, src, png, relay);
        entry.lastPushedAt = new Date().toISOString();
        summary.pushed += 1;
      } catch {
        summary.failed += 1;
      }
    }
    writeRegistry(context, registry);
    await context.sync();
    return summary;
  });
}

export async function goToSource(id: string): Promise<void> {
  await Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const entry = entryOf(registry, id, "go to source");
    const resolved = await resolveSource(context, entry);
    if (resolved === null) {
      throw new Error(`go to source ${entry.label}: source missing`);
    }
    if (resolved.kind === "range") {
      resolved.range.select();
    } else {
      resolved.chart.worksheet.activate();
      resolved.chart.activate();
    }
    await context.sync();
  });
}

export async function removeLink(id: string, relay: RelayApi): Promise<void> {
  await Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const entry = entryOf(registry, id, "remove");
    await releaseAnchor(context, entry);
    registry.links = registry.links.filter((link) => link.id !== id);
    writeRegistry(context, registry);
    await context.sync();
    await forget(entry, relay);
  });
}
