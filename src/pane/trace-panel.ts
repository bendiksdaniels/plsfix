// Audit overlay toggle and Smart Track: walk precedents/dependents one hop at
// a time with a back stack, or paint the workbook's audit tint. Office.js
// only reaches here through ../excel; the guard and its busy state are
// shared with every other tab through ./shared.

import {
  lastAuditNote,
  parseAddress,
  selectArea,
  toggleAuditOverlay,
  traceActiveCell,
  type TraceArea,
  type TraceDirection,
  type TraceResult,
} from "../excel";
import { getElement } from "../ui/dom";
import { guard } from "./shared";

const TRACE_STACK_LIMIT = 20;

interface TraceView {
  direction: TraceDirection;
  result: TraceResult;
}

const traceStack: TraceView[] = [];
let traceView: TraceView | null = null;
let auditOn = false;

export function renderAuditState(): void {
  getElement("audit-state").textContent = auditOn ? "On" : "Off";
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

  const { direction, result } = traceView;
  getElement("trace-origin").textContent = `${result.origin} · ${direction}`;
  getElement<HTMLButtonElement>("trace-back").disabled =
    traceStack.length === 0;

  if (result.areas.length === 0) {
    const empty = document.createElement("p");
    empty.className = "hint";
    empty.textContent = `No direct ${direction}.`;
    chips.append(empty);
  }

  for (const area of result.areas) {
    const label = `${area.sheet}!${area.address}`;
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "chip";
    chip.textContent = label;
    chip.title = `Select ${label} (${area.cellCount.toLocaleString()} cells)`;
    chip.addEventListener("click", () => void guard(() => walkTo(area)));
    chips.append(chip);
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
  const first = result.areas[0];
  if (jump && first) await selectArea(first);

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
  // Skipped protected cells, if any, ride along in the same toast.
  const note = lastAuditNote();
  return note ? `${state}: ${note}` : state;
}
