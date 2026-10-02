// Property-fuzz over the pure undo ring (pushCapped): many random sequences
// of capture/commit/discard/undo, shaped like undo.ts's own state machine
// but built only on the exported pure function, checked after every single
// step for: never restores a discarded or evicted capture, evicts the
// OLDEST entries first (never a newer one while an older one survives),
// both limits (five slots, the 25 000-cell budget) always hold, and the
// "N more to undo" count always matches the real stack length after a pop.

import { describe, expect, it } from "vitest";
import { pushCapped, type Weighted } from "../../src/excel/undo-stack";
import { UNDO_CELL_BUDGET, UNDO_DEPTH } from "../../src/excel/undo";

interface Entry extends Weighted {
  id: number;
}

const LIMITS = { maxDepth: UNDO_DEPTH, maxCells: UNDO_CELL_BUDGET };

// Deterministic PRNG (mulberry32): reproducible, no external dependency.
function mulberry32(seed: number): () => number {
  let a = seed;
  return () => {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function totalCells(stack: readonly Entry[]): number {
  return stack.reduce((sum, entry) => sum + entry.cells, 0);
}

// Mirrors undo.ts's own rules exactly (captureUndoAreas/commitUndo/discardUndo/
// undoLastAction), built only on the real pushCapped: a capture discards
// whatever was still pending first, an undo does the same before popping.
class UndoModel {
  stack: Entry[] = [];
  private pending: Entry | null = null;
  private nextId = 0;
  readonly everGone = new Set<number>();

  private forget(gone: readonly Entry[]): void {
    for (const entry of gone) this.everGone.add(entry.id);
  }

  private discardPending(): void {
    if (this.pending && this.stack[0]?.id === this.pending.id) {
      this.forget([this.stack[0]]);
      this.stack = this.stack.slice(1);
    }
    this.pending = null;
  }

  // Weight kept well under UNDO_CELL_BUDGET on its own (SELECTION_CELL_CAP is
  // 5 000, a fifth of the budget in real life) - undo.ts never lets a single
  // capture past that cap reach pushCapped at all.
  capture(cells: number): void {
    this.discardPending();
    const before = this.stack;
    const entry: Entry = { id: this.nextId, cells };
    this.nextId += 1;
    this.stack = pushCapped(before, entry, LIMITS);

    // Evicts oldest first: the survivors from the old stack must be exactly
    // its own front (newest) prefix, in order - never a gap, never a swap.
    const oldSurvivors = this.stack.filter((item) => item.id !== entry.id);
    expect(oldSurvivors).toEqual(before.slice(0, oldSurvivors.length));
    this.forget(before.slice(oldSurvivors.length));

    this.pending = this.stack[0]?.id === entry.id ? entry : null;
    // A capture within the realistic weight range must always survive its
    // own push (only a pathological single entry over budget would not).
    expect(this.stack[0]?.id).toBe(entry.id);
  }

  commit(): void {
    this.pending = null;
  }

  discard(): void {
    this.discardPending();
  }

  // Returns the "N more to undo" count, or null when there was nothing to undo.
  undo(): number | null {
    this.discardPending();
    if (this.stack.length === 0) return null;
    this.forget([this.stack[0]!]);
    this.stack = this.stack.slice(1);
    return this.stack.length;
  }

  assertInvariants(): void {
    expect(this.stack.length).toBeLessThanOrEqual(UNDO_DEPTH);
    expect(totalCells(this.stack)).toBeLessThanOrEqual(UNDO_CELL_BUDGET);
    const ids = this.stack.map((entry) => entry.id);
    expect(new Set(ids).size).toBe(ids.length); // no duplicates
    for (let i = 1; i < ids.length; i += 1) {
      expect(ids[i]).toBeLessThan(ids[i - 1]!); // strictly newest-first
    }
    for (const entry of this.stack) {
      expect(this.everGone.has(entry.id)).toBe(false); // nothing gone is back
    }
  }
}

const OPS = [
  "capture",
  "capture",
  "capture",
  "commit",
  "discard",
  "undo",
] as const;

function runSequence(seed: number, steps: number): void {
  const random = mulberry32(seed);
  const model = new UndoModel();

  for (let step = 0; step < steps; step += 1) {
    const op = OPS[Math.floor(random() * OPS.length)]!;
    if (op === "capture") {
      // 1..5000: SELECTION_CELL_CAP is the real ceiling captureUndoAreas ever
      // lets through to pushCapped.
      const cells = 1 + Math.floor(random() * 5_000);
      model.capture(cells);
    } else if (op === "commit") {
      model.commit();
    } else if (op === "discard") {
      model.discard();
    } else {
      const remaining = model.undo();
      if (remaining !== null) {
        expect(remaining).toBe(model.stack.length);
      }
    }
    model.assertInvariants();
  }
}

describe("pls,fix Undo's stack under random sequences", () => {
  it.each([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])(
    "holds every invariant over 500 random steps (seed %i)",
    (seed) => {
      runSequence(seed * 104_729, 500);
    },
  );

  it("evicts oldest first when a burst of full-size captures overflows depth", () => {
    const model = new UndoModel();
    for (let i = 0; i < UNDO_DEPTH + 3; i += 1) {
      model.capture(1);
      model.commit();
    }
    expect(model.stack).toHaveLength(UNDO_DEPTH);
    // The three oldest (ids 0, 1, 2) must be gone; the five newest survive.
    for (let id = 0; id < 3; id += 1) expect(model.everGone.has(id)).toBe(true);
    expect(model.stack.map((entry) => entry.id)).toEqual([7, 6, 5, 4, 3]);
  });

  it("evicts oldest first when captures overflow the cell budget alone", () => {
    const model = new UndoModel();
    // Five captures of 5 000 cells sit right at the depth limit (5) and well
    // under budget (25 000); a sixth of 5 000 pushes the total to 30 000,
    // forcing an eviction on budget even though depth alone would allow six.
    for (let i = 0; i < 5; i += 1) {
      model.capture(5_000);
      model.commit();
    }
    expect(model.stack).toHaveLength(5);
    model.capture(5_000);
    model.commit();
    expect(model.stack.length).toBeLessThan(6);
    expect(totalCells(model.stack)).toBeLessThanOrEqual(UNDO_CELL_BUDGET);
    expect(model.everGone.has(0)).toBe(true); // the very first capture is gone
    expect(model.stack[0]!.id).toBe(5); // the newest just pushed
  });

  it("never lets a discarded (refused-write) capture come back on undo", () => {
    const model = new UndoModel();
    model.capture(10);
    model.commit(); // entry 0 is now permanent
    model.capture(20);
    model.discard(); // entry 1's write was refused: dropped, never offered
    expect(model.stack.map((entry) => entry.id)).toEqual([0]);
    expect(model.everGone.has(1)).toBe(true);
    expect(model.undo()).toBe(0);
    expect(model.everGone.has(0)).toBe(true);
    expect(model.undo()).toBeNull(); // nothing left, never entry 1 either
  });

  it("a fresh capture drops a still-pending one before pushing (never two pending)", () => {
    const model = new UndoModel();
    model.capture(5); // entry 0, pending
    model.capture(7); // its own flow abandoned entry 0's write mid-air
    expect(model.stack.map((entry) => entry.id)).toEqual([1]);
    expect(model.everGone.has(0)).toBe(true);
  });
});
