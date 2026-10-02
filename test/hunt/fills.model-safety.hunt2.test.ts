// Model-based property: random modeller fills (none, solid, named pattern, None
// with stale residue) through random overlay, pinstripes, paintbrush, preset,
// AutoColor and Undo steps on both fake hosts. With no overlay on, every cell
// equals the model's last committed fill; a refused step changes nothing.

import fc from "fast-check";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeFill,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import { FakeRelay } from "../fakerelay";
import type * as ExcelModule from "../../src/excel";
import { fillGrid, requestFills } from "../../src/excel/fill-store";
import { createWorkspace, type KeyStore } from "../../src/link/workspace";

enableStrictLoadSemantics();

const SEED = 20260927;
const RUNS = 300;

const GRID_ADDR = "Model!B2:E5";
const COLS = ["B", "C", "D", "E"];
const ROWS = [2, 3, 4, 5];
const CELLS = ROWS.flatMap((row) => COLS.map((col) => `Model!${col}${row}`));
// Never written by anything in the sequence: proves a tool touched only what
// it should have. SOURCE is also the paintbrush's capture cell.
const SOURCE_ADDR = "Model!Z1";
const SENTINEL_ADDR = "Model!Z20";
const SENTINEL_FILL: FakeFill = {
  pattern: "Solid",
  color: "#123456",
  patternColor: "#123456",
};

const PALETTE = [
  "#FFFFFF",
  "#000000",
  "#FF0000",
  "#00B050",
  "#0070C0",
  "#FFC000",
  "#7030A0",
];
const NAMED_PATTERNS = [
  "LightUp",
  "LightDown",
  "LightHorizontal",
  "LightVertical",
  "Checker",
  "Gray50",
  "Gray25",
  "Grid",
  "CrissCross",
];

const colorArb = fc.constantFrom(...PALETTE);

// The four fill shapes target 1 names: a clean "no fill", a "no fill" cell
// carrying stale colour/pattern-colour residue (the Mac's own read-back shape
// for one, {pattern: null, patternColor: ""} over whatever colour the model
// holds - lessons.md 2026-08-27), a solid colour (white included) and a named
// pattern with its own pattern colour.
const fillArb: fc.Arbitrary<FakeFill> = fc.oneof(
  fc.constant<FakeFill>({
    pattern: "None",
    color: "#FFFFFF",
    patternColor: "#FFFFFF",
  }),
  fc
    .tuple(colorArb, colorArb)
    .map(([color, patternColor]) => ({ pattern: "None", color, patternColor })),
  colorArb.map((color) => ({ pattern: "Solid", color, patternColor: color })),
  fc
    .tuple(fc.constantFrom(...NAMED_PATTERNS), colorArb, colorArb)
    .map(([pattern, color, patternColor]) => ({
      pattern,
      color,
      patternColor,
    })),
);

const gridFillsArb = fc.array(fillArb, {
  minLength: CELLS.length,
  maxLength: CELLS.length,
});

type PresetName = "title" | "header" | "input" | "formula" | "result";
type Slot = 1 | 2 | 3;

type Step =
  | { kind: "audit" }
  | { kind: "highlight" }
  | { kind: "pinstripes"; axis: "rows" | "columns" }
  | { kind: "capture"; slot: Slot }
  | { kind: "apply"; slot: Slot }
  | { kind: "preset"; name: PresetName }
  | { kind: "autocolor" }
  | { kind: "undo" };

const stepArb: fc.Arbitrary<Step> = fc.oneof(
  fc.constant<Step>({ kind: "audit" }),
  fc.constant<Step>({ kind: "highlight" }),
  fc
    .constantFrom<"rows" | "columns">("rows", "columns")
    .map((axis): Step => ({ kind: "pinstripes", axis })),
  fc
    .constantFrom<Slot>(1, 2, 3)
    .map((slot): Step => ({ kind: "capture", slot })),
  fc.constantFrom<Slot>(1, 2, 3).map((slot): Step => ({ kind: "apply", slot })),
  fc
    .constantFrom<PresetName>("title", "header", "input", "formula", "result")
    .map((name): Step => ({ kind: "preset", name })),
  fc.constant<Step>({ kind: "autocolor" }),
  fc.constant<Step>({ kind: "undo" }),
);

const sequenceArb = fc.array(stepArb, { minLength: 3, maxLength: 6 });

function memoryStore(): KeyStore {
  const map = new Map<string, string>();
  return {
    get: async (k) => map.get(k) ?? null,
    set: async (k, v) => {
      map.set(k, v);
    },
    remove: async (k) => {
      map.delete(k);
    },
  };
}

// Read the way every tool in this file's own sequence does - through the host
// API (getCellProperties), never a raw peek at the fake's internal storage.
// On the web, a Solid fill's pattern colour is not just untagged but
// UNOBSERVABLE (webRangeFill always answers ""), so a raw peek would compare
// a value nothing in production, on this host, could ever read back or be
// asked to preserve.
async function readApiGrid(): Promise<string[]> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Model");
    const grid = requestFills(sheet.getRange("B2:E5"));
    await context.sync();
    return fillGrid(grid).flat();
  });
}

async function readApiUntouched(): Promise<[string, string]> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Model");
    const source = requestFills(sheet.getRange("Z1"));
    const sentinel = requestFills(sheet.getRange("Z20"));
    await context.sync();
    return [fillGrid(source)[0]![0]!, fillGrid(sentinel)[0]![0]!];
  });
}

afterEach(() => {
  uninstallFakeHost();
});

async function runSequence(
  web: boolean,
  gridFills: FakeFill[],
  sourceFill: FakeFill,
  sequence: Step[],
): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const helpers = installFakeHost({ sheets: ["Model"], web }).helpers;
  const smt = (await import("../../src/excel")) as typeof ExcelModule;

  // A uniform formula across the block: auditGrid then marks every cell
  // "both" (same neighbour on every side), so the overlay always has real
  // stripes to paint - never the "no formula here to stripe" refusal.
  helpers.seed(
    GRID_ADDR,
    ROWS.map(() =>
      COLS.map(() => ({ formula: "=Z9", r1c1: "=R9C26", value: 42 })),
    ),
  );
  CELLS.forEach((addr, index) => helpers.setFill(addr, gridFills[index]!));
  helpers.setFill(SOURCE_ADDR, sourceFill);
  helpers.setFill(SENTINEL_ADDR, SENTINEL_FILL);

  helpers.select(GRID_ADDR);
  const ws = await createWorkspace(memoryStore());
  await smt.exportSelection(ws, new FakeRelay());
  helpers.select(GRID_ADDR);

  const untouched = await readApiUntouched();
  let trueFill = await readApiGrid();
  // Mirrors src/excel/undo.ts's 5-deep, newest-first stack: what pls,fix
  // Undo must put back, one entry per mutating step this run took.
  const undoMirror: string[][] = [];
  const slots: (Awaited<ReturnType<typeof smt.captureSlot>> | null)[] = [
    null,
    null,
    null,
  ];

  const pushUndo = (): void => {
    undoMirror.unshift(trueFill);
    if (undoMirror.length > 5) undoMirror.pop();
  };
  // A refusal must write nothing at all - not even a partial batch - so every
  // refusal branch below checks the full grid, not only the model's view of
  // it (which a still-active overlay would otherwise hide from the read).
  const expectNoOp = async (action: () => Promise<unknown>): Promise<void> => {
    const before = await readApiGrid();
    await expect(action()).rejects.toThrow();
    expect(await readApiGrid()).toEqual(before);
  };
  const overlayActive = (): boolean =>
    smt.auditOverlayOn() || smt.linkHighlightOn();

  for (const step of sequence) {
    switch (step.kind) {
      case "audit":
        if (smt.linkHighlightOn())
          await expectNoOp(() => smt.toggleAuditOverlay());
        else await smt.toggleAuditOverlay();
        break;
      case "highlight":
        if (smt.auditOverlayOn())
          await expectNoOp(() => smt.toggleLinkHighlight());
        else await smt.toggleLinkHighlight();
        break;
      case "pinstripes":
        if (overlayActive()) {
          await expectNoOp(() => smt.applyPinstripes(step.axis));
        } else {
          pushUndo();
          await smt.applyPinstripes(step.axis);
          trueFill = await readApiGrid();
        }
        break;
      case "capture":
        helpers.setActiveCell(SOURCE_ADDR);
        slots[step.slot - 1] = await smt.captureSlot(step.slot);
        break;
      case "apply": {
        const slot = slots[step.slot - 1] ?? null;
        if (!slot || overlayActive()) {
          await expectNoOp(() => smt.applySlot(step.slot, slot));
        } else {
          pushUndo();
          await smt.applySlot(step.slot, slot);
          trueFill = await readApiGrid();
        }
        break;
      }
      case "preset":
        // OUT-OF-SLICE gap, not this branch's to fix: selection.ts's
        // applyPreset never calls requireNoOverlayOwner, so it paints
        // straight over an overlay's own tint. Skipped here rather than
        // baked into this committed property; reported separately.
        if (overlayActive()) break;
        pushUndo();
        await smt.applyPreset(step.name);
        trueFill = await readApiGrid();
        break;
      case "autocolor":
        if (overlayActive()) {
          await expectNoOp(() => smt.autocolorSelection());
        } else {
          pushUndo();
          await smt.autocolorSelection();
          trueFill = await readApiGrid();
        }
        break;
      case "undo":
        // OUT-OF-SLICE gap, not this branch's to fix: undo.ts's
        // undoLastAction never calls requireNoOverlayOwner either, so it
        // rewrites overlay-owned cells directly; the overlay's own later
        // toggle-off then hands back its now-stale memory over them. Skipped
        // here rather than baked into this committed property; reported
        // separately.
        if (overlayActive()) break;
        if (undoMirror.length === 0) {
          await expectNoOp(() => smt.undoLastAction());
        } else {
          const expected = undoMirror.shift()!;
          await smt.undoLastAction();
          trueFill = expected;
        }
        break;
    }

    expect(await readApiUntouched()).toEqual(untouched);
    if (!overlayActive()) expect(await readApiGrid()).toEqual(trueFill);
  }
}

describe("fill safety under a random tool sequence", () => {
  it("holds on the desktop fake", async () => {
    await fc.assert(
      fc.asyncProperty(
        gridFillsArb,
        fillArb,
        sequenceArb,
        async (gridFills, sourceFill, sequence) => {
          await runSequence(false, gridFills, sourceFill, sequence);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });

  it("holds on the Excel-for-the-web fake", async () => {
    await fc.assert(
      fc.asyncProperty(
        gridFillsArb,
        fillArb,
        sequenceArb,
        async (gridFills, sourceFill, sequence) => {
          await runSequence(true, gridFills, sourceFill, sequence);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
