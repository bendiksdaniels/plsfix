// A pure ring for pls,fix Undo's stack: newest first, bounded by both an
// entry count and a total cell-weight budget. No Office.js and nothing
// undo.ts-specific, so the eviction rule is testable on its own instead of
// only at the point where undo.ts's real depth and budget numbers coincide.

export interface Weighted {
  cells: number;
}

export interface StackLimits {
  maxDepth: number;
  maxCells: number;
}

// Adds `entry` to the front of `stack` (newest first) and drops from the back
// (oldest) until both limits hold. `stack` is read-only and never mutated;
// the caller just reassigns its own `let` binding to the result.
export function pushCapped<T extends Weighted>(
  stack: readonly T[],
  entry: T,
  limits: StackLimits,
): T[] {
  const next = [entry, ...stack];
  while (next.length > limits.maxDepth || totalCells(next) > limits.maxCells) {
    next.pop();
  }
  return next;
}

function totalCells(stack: readonly Weighted[]): number {
  return stack.reduce((sum, entry) => sum + entry.cells, 0);
}
