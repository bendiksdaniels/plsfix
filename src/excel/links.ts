// The Excel side of tracked links, one function per pane action: export a
// selection or the active chart, push every anchored source again, list what
// this workbook owns, jump back to a source and remove a link. The anchor - not
// the address it was created at - is what every later flow resolves through.

import { randomBytes } from "../link/crypto";
import {
  anchorName,
  newLinkId,
  sourceLabel,
  type RegistryEntry,
} from "../link/model";
import type { RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import { SELECTION_CELL_CAP, selectedSingleRange } from "./internal";
import {
  createChartAnchor,
  createRangeAnchor,
  entryOf,
  forget,
  newEntry,
  readRegistry,
  refuseAnchoredChart,
  releaseAnchor,
  renderAnchored,
  renderSource,
  requireImageApi,
  resolveSource,
  sourceOf,
  workbookName,
  writeRegistry,
  type ResolvedSource,
} from "./link-anchors";
import { publish, pushPayload, type NewLink } from "./link-record";
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
  // One "<label>: <reason>" per failed push, so the pane can say why rather
  // than only how many.
  failures: string[];
}

export async function exportSelection(
  ws: Workspace,
  relay: RelayApi,
): Promise<ExportResult> {
  requireImageApi();
  const workbook = await workbookName();
  return Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const range = await selectedSingleRange(context, "export");
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
    const id = newLinkId(randomBytes);
    const anchor = anchorName(id);
    const src = sourceOf(workbook, anchor, resolved);
    const entry = newEntry(id, "range", anchor, sourceLabel(src, "range"));
    // Anchor and picture in one batch, before any network call: a selection
    // that changes during the upload cannot make the two describe different
    // objects. The render owns the anchor from here on, so a picture that
    // never arrives takes the name with it.
    const named = createRangeAnchor(context, range, anchor);
    const release = () => named.delete();
    const png = await renderAnchored(context, resolved, entry.label, release);

    const link: NewLink = { entry, src, png, registry, release };
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
  if (typeof callable !== "function") {
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
  requireImageApi();
  const workbook = await workbookName();
  return Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const chart = await activeChart(context);
    const previousName = chart.name;
    refuseAnchoredChart(registry, previousName);

    const resolved: ResolvedSource = {
      kind: "chart",
      sheet: chart.worksheet.name,
      ref: previousName,
      chart,
      width: chart.width,
      height: chart.height,
    };
    const id = newLinkId(randomBytes);
    const anchor = anchorName(id);
    const src = sourceOf(workbook, anchor, resolved);
    const entry = newEntry(id, "chart", anchor, sourceLabel(src, "chart"));
    // The rename commits with the batch that asks for the picture, so a render
    // that fails has to give the chart its own name back: an SMT_LINK_ chart no
    // registry entry claims is one the modeller cannot export again.
    createChartAnchor(chart, anchor);
    const release = () => createChartAnchor(chart, previousName);
    const png = await renderAnchored(context, resolved, entry.label, release);

    const link: NewLink = { entry, src, png, registry, release };
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

    const summary: PushSummary = {
      pushed: 0,
      missing: 0,
      failed: 0,
      failures: [],
    };
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
      } catch (error) {
        const reason = error instanceof Error ? error.message : String(error);
        summary.failed += 1;
        summary.failures.push(`${entry.label}: ${reason}`);
      }
    }
    writeRegistry(context, registry);
    await context.sync();
    return summary;
  });
}

// Anchors are searched on every worksheet, hidden ones included, but Excel
// refuses to activate a sheet that is hidden: a linked range on a hidden calc
// sheet would fail the sync with a bare InvalidOperation.
async function requireVisibleSheet(
  context: Excel.RequestContext,
  resolved: ResolvedSource,
  label: string,
): Promise<void> {
  const sheet =
    resolved.kind === "range"
      ? resolved.range.worksheet
      : resolved.chart.worksheet;
  sheet.load("name,visibility");
  await context.sync();
  if (sheet.visibility !== Excel.SheetVisibility.visible) {
    throw new Error(`go to source ${label}: sheet "${sheet.name}" is hidden`);
  }
}

export async function goToSource(id: string): Promise<void> {
  await Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const entry = entryOf(registry, id, "go to source");
    const resolved = await resolveSource(context, entry);
    if (resolved === null) {
      throw new Error(`go to source ${entry.label}: source missing`);
    }
    await requireVisibleSheet(context, resolved, entry.label);
    if (resolved.kind === "range") {
      // Excel refuses to select on a sheet that is not the active one.
      resolved.range.worksheet.activate();
      resolved.range.select();
    } else {
      resolved.chart.worksheet.activate();
      resolved.chart.activate();
    }
    await context.sync();
  });
}

// The relay copy is revoked first. A revoke that fails throws before anything
// local moves, so the entry - and the token the revoke needs - is still there
// to try again with; the other order leaves the relay serving a picture the
// workbook no longer knows how to withdraw.
export async function removeLink(id: string, relay: RelayApi): Promise<void> {
  await Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const entry = entryOf(registry, id, "remove");
    await forget(entry, relay);
    await releaseAnchor(context, entry);
    registry.links = registry.links.filter((link) => link.id !== id);
    writeRegistry(context, registry);
    await context.sync();
  });
}
