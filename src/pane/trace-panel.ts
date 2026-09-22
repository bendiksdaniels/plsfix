// Audit overlay toggle and Smart Track: walk precedents/dependents one hop at
// a time with a back stack, or paint the workbook's audit tint. Office.js
// only reaches here through ../excel; the guard and its busy state are
// shared with every other tab through ./shared.

import {
  lastAuditNote,
  parseAddress,
  selectArea,
  selectAreas,
  toggleAuditOverlay,
  traceActiveCell,
  tracePrecedentsOfSelection,
  type CellPrecedents,
  type MultiTraceResult,
  type TraceArea,
  type TraceDirection,
  type TraceResult,
} from "../excel";
import { getElement } from "../ui/dom";
import { guard } from "./shared";

const TRACE_STACK_LIMIT = 20;
const NO_FORMULAS = "No formulas in the selection.";

interface TraceView {
  direction: TraceDirection;
  result: TraceResult;
  /** Set only by "Precedents of selection": the list is per source cell. */
  groups?: CellPrecedents[];
}

const traceStack: TraceView[] = [];
let traceView: TraceView | null = null;
let auditOn = false;

export function renderAuditState(): void {
  getElement("audit-state").textContent = auditOn ? "On" : "Off";
}

function chipFor(area: TraceArea): HTMLButtonElement {
  const label = `${area.sheet}!${area.address}`;
  const chip = document.createElement("button");
  chip.type = "button";
  chip.className = "chip";
  chip.textContent = label;
  chip.title = `Select ${label} (${area.cellCount.toLocaleString()} cells)`;
  chip.addEventListener("click", () => void guard(() => walkTo(area)));
  return chip;
}

// The chip strip wraps, and .chip-group takes a whole row of it: that is what
// makes the chips under a label read as belonging to it.
function lineFor(text: string): HTMLParagraphElement {
  const line = document.createElement("p");
  line.className = "hint chip-group";
  line.textContent = text;
  return line;
}

// One label per source cell, its own chips under it. A source cell that reads
// from nothing still gets its line: it is why the count is lower than the
// number of cells asked about.
function renderGroups(chips: HTMLElement, groups: CellPrecedents[]): void {
  if (groups.length === 0) {
    chips.append(lineFor(NO_FORMULAS));
    return;
  }
  for (const group of groups) {
    const empty = group.areas.length === 0;
    chips.append(lineFor(empty ? `${group.cell}: no precedents` : group.cell));
    for (const area of group.areas) chips.append(chipFor(area));
  }
}

// Back only ever has something to go to once a walk has pushed onto the
// stack, which itself only happens with a view already showing - so this is
// safe to call any time, including to put the button back after setBusy's
// blanket disable, whether or not a trace is on screen right now.
export function syncTraceBack(): void {
  getElement<HTMLButtonElement>("trace-back").disabled =
    traceStack.length === 0;
}

function renderTrace(): void {
  const panel = getElement<HTMLDivElement>("trace-panel");
  const chips = getElement<HTMLDivElement>("trace-chips");
  // Clearing before the early return keeps stale chips from firing when the
  // panel comes back.
  chips.replaceChildren();

  if (!traceView) {
    panel.hidden = true;
    return;
  }

  const { direction, result, groups } = traceView;
  getElement("trace-origin").textContent = `${result.origin} · ${direction}`;
  syncTraceBack();

  if (groups) {
    renderGroups(chips, groups);
  } else {
    if (result.areas.length === 0) {
      chips.append(lineFor(`No direct ${direction}.`));
    }
    for (const area of result.areas) chips.append(chipFor(area));
  }

  panel.hidden = false;
}

function traceMessage(direction: TraceDirection, count: number): string {
  if (count === 0) return `No direct ${direction}`;
  return `${count} direct ${count === 1 ? direction.slice(0, -1) : direction}`;
}

async function showTrace(
  direction: TraceDirection,
  jump: boolean,
): Promise<string> {
  const result = await traceActiveCell(direction);
  traceView = { direction, result };
  renderTrace();

  // A shortcut can fire with the pane closed, so the keystroke jumps instead.
  // The jump selects every area on the first one's sheet at once; an area on
  // another sheet stays a chip, same as it always has.
  if (jump && result.areas.length > 0) await selectAreas(result.areas);

  return traceMessage(direction, result.areas.length);
}

// A fresh trace is a new walk, so the back stack starts empty.
export function startTrace(
  direction: TraceDirection,
  jump: boolean,
): Promise<string> {
  traceStack.length = 0;
  return showTrace(direction, jump);
}

function plural(count: number, one: string, many: string): string {
  return `${count.toLocaleString()} ${count === 1 ? one : many}`;
}

// "18 precedents of 6 cells; 2 cells have no formula" - what was found, over
// how many cells, and why the rest of the selection was left out.
function precedentsMessage(result: MultiTraceResult): string {
  if (result.formulaCells === 0) return NO_FORMULAS;
  const cells = plural(result.formulaCells, "cell", "cells");
  const found =
    result.areas.length === 0
      ? `No precedents of ${cells}`
      : `${plural(result.areas.length, "precedent", "precedents")} of ${cells}`;
  if (result.skipped === 0) return found;
  return `${found}; ${plural(result.skipped, "cell has", "cells have")} no formula`;
}

/**
 * "Precedents of selection": every formula cell of the selection at once. The
 * list is grouped by source cell, and a chip carries on as a one-cell walk, so
 * the back stack starts empty exactly as a fresh single-cell trace does.
 */
export async function startPrecedentsOfSelection(): Promise<string> {
  const result = await tracePrecedentsOfSelection();
  traceStack.length = 0;
  traceView = {
    direction: "precedents",
    result: { origin: result.origin, areas: result.areas },
    groups: result.groups,
  };
  renderTrace();
  return precedentsMessage(result);
}

async function walkTo(area: TraceArea): Promise<string> {
  const current = traceView;
  if (!current) return "Nothing to trace";

  await selectArea(area);
  if (traceStack.length >= TRACE_STACK_LIMIT) traceStack.shift();
  traceStack.push(current);
  return showTrace(current.direction, false);
}

export async function traceBack(): Promise<string> {
  const previous = traceStack.pop();
  if (!previous) return "Nothing to go back to";

  traceView = previous;
  renderTrace();
  await selectArea(parseAddress(previous.result.origin));
  return `Back at ${previous.result.origin}`;
}

export async function toggleAudit(): Promise<string> {
  auditOn = await toggleAuditOverlay();
  renderAuditState();
  const state = auditOn ? "Audit overlay on" : "Audit overlay off";
  // The adapter's note is the bare sentence, with no stage of its own
  // (src/excel/protection.ts's protectedSentence) - composed here so the
  // stage is named exactly once, never doubled and never dropped.
  const note = lastAuditNote();
  return note ? `${state}: ${note}` : state;
}
