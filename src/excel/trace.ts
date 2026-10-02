// Smart Track: direct precedents and dependents of the active cell, gated on the
// host's Excel API set (older Excel builds lack the method entirely).

import { hostSupports } from "./internal";
import { parseAddress } from "./shared";

export type TraceDirection = "precedents" | "dependents";

export interface TraceArea {
  sheet: string;
  address: string;
  cellCount: number;
}

export interface TraceResult {
  origin: string;
  areas: TraceArea[];
}

const TRACE_API_SET: Record<TraceDirection, string> = {
  precedents: "1.12",
  dependents: "1.13",
};

/**
 * Queues one direct-trace call on a cell, with the API guard in front of it
 * and the load the answer needs. The batch the caller syncs decides when it
 * answers, so several cells can be asked in one round trip - which is what
 * "Precedents of selection" does (src/excel/trace-many.ts).
 */
export function queueTrace(
  cell: Excel.Range,
  direction: TraceDirection,
): Excel.WorkbookRangeAreas {
  // The hosted office.js always defines the method, so the host API set decides.
  const method =
    direction === "precedents" ? "getDirectPrecedents" : "getDirectDependents";
  const callable = (cell as unknown as Record<string, unknown>)[method];
  if (
    typeof callable !== "function" ||
    !hostSupports(TRACE_API_SET[direction])
  ) {
    throw new Error("Tracing needs a newer Excel build.");
  }

  const found =
    direction === "precedents"
      ? cell.getDirectPrecedents()
      : cell.getDirectDependents();
  found.ranges.load("items/address,items/cellCount");
  return found;
}

/** What a synced trace call holds, one area per rectangle Excel named. */
export function traceAreas(found: Excel.WorkbookRangeAreas): TraceArea[] {
  return found.ranges.items.map((item) => ({
    ...parseAddress(item.address),
    cellCount: item.cellCount,
  }));
}

export async function traceActiveCell(
  direction: TraceDirection,
): Promise<TraceResult> {
  return Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.load("address");
    await context.sync();

    const found = queueTrace(cell, direction);
    try {
      await context.sync();
    } catch (error) {
      // Excel reports "nothing found" by throwing rather than returning nothing.
      if ((error as { code?: string }).code !== Excel.ErrorCodes.itemNotFound) {
        throw error;
      }
      return { origin: cell.address, areas: [] };
    }

    return { origin: cell.address, areas: traceAreas(found) };
  });
}

export async function selectArea(area: {
  sheet: string;
  address: string;
}): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = area.sheet
      ? context.workbook.worksheets.getItem(area.sheet)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.activate();
    sheet.getRange(area.address).select();
    await context.sync();
  });
}

/**
 * Every area on the FIRST area's sheet, selected together as one multi-area
 * selection, that sheet activated. A precedent or dependent on another sheet
 * is left for the panel's chips exactly as it is today: jumping there too
 * would leave the modeller looking at a cell on a sheet they did not ask for.
 */
export async function selectAreas(areas: TraceArea[]): Promise<void> {
  const first = areas[0];
  if (!first) return;
  const onFirstSheet = areas.filter((area) => area.sheet === first.sheet);

  await Excel.run(async (context) => {
    const sheet = first.sheet
      ? context.workbook.worksheets.getItem(first.sheet)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.activate();
    sheet
      .getRanges(onFirstSheet.map((area) => area.address).join(","))
      .select();
    await context.sync();
  });
}
