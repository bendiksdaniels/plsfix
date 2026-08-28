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

export async function traceActiveCell(
  direction: TraceDirection,
): Promise<TraceResult> {
  return Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.load("address");
    await context.sync();

    // The hosted office.js always defines the method, so the host API set decides.
    const method =
      direction === "precedents"
        ? "getDirectPrecedents"
        : "getDirectDependents";
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

    try {
      await context.sync();
    } catch (error) {
      // Excel reports "nothing found" by throwing rather than returning nothing.
      if ((error as { code?: string }).code !== Excel.ErrorCodes.itemNotFound) {
        throw error;
      }
      return { origin: cell.address, areas: [] };
    }

    return {
      origin: cell.address,
      areas: found.ranges.items.map((item) => ({
        ...parseAddress(item.address),
        cellCount: item.cellCount,
      })),
    };
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
