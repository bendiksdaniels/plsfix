// Variance reconciliation: read the numeric cells in one selected block, find
// a subset that reaches the requested target, then leave those cells selected
// so the modeller can inspect the answer in place.

import { RECONCILE_MAX_VALUES, solveReconciliation } from "../reconcile";
import { selectedSingleRange, withinCap } from "./internal";
import { parseAddress } from "./shared";

export interface ReconciliationSelection {
  addresses: string[];
  count: number;
  difference: number;
  sum: number;
  values: number[];
}

interface NumericCell {
  column: number;
  row: number;
  value: number;
}

export async function reconcileSelection(
  target: number,
  tolerance: number,
): Promise<ReconciliationSelection> {
  return Excel.run(async (context) => {
    // The cap is checked before the values are asked for: a clicked column
    // header is a million cells, and the 34-cell rule only runs after the read.
    const range = await withinCap(
      context,
      await selectedSingleRange(context, "Reconciliation"),
      "Reconciliation",
    );
    range.load("values,rowCount,columnCount");
    await context.sync();

    const cells: NumericCell[] = [];
    for (let row = 0; row < range.rowCount; row += 1) {
      for (let column = 0; column < range.columnCount; column += 1) {
        const value = range.values[row]?.[column];
        if (typeof value === "number" && Number.isFinite(value)) {
          cells.push({ row, column, value });
        }
      }
    }

    if (cells.length === 0) {
      throw new Error("Select a range containing numeric cells.");
    }
    if (cells.length > RECONCILE_MAX_VALUES) {
      throw new Error(
        `Reconciliation supports up to ${String(RECONCILE_MAX_VALUES)} numeric cells; select a smaller range.`,
      );
    }

    const match = solveReconciliation(
      cells.map((cell) => cell.value),
      target,
      tolerance,
    );
    if (match === null) {
      throw new Error(
        "No combination reaches the target within the tolerance.",
      );
    }

    const matchedRanges = match.indices.map((index) => {
      const cell = cells[index]!;
      const matched = range.getCell(cell.row, cell.column);
      matched.load("address");
      return matched;
    });
    await context.sync();

    const addresses = matchedRanges.map(
      (matched) => parseAddress(matched.address).address,
    );
    range.worksheet.getRanges(addresses.join(",")).select();
    await context.sync();

    return {
      addresses,
      count: addresses.length,
      difference: match.difference,
      sum: match.sum,
      values: match.indices.map((index) => cells[index]!.value),
    };
  });
}
