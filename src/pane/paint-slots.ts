// Paintbrush slots: three captured formats, kept on the machine like the
// brand palette. They outlive the pane, and nothing about them is written
// into the workbook. Office.js only reaches here through ../excel.

import { applySlot, captureSlot } from "../excel";
import {
  emptySlots,
  type PaintSlots,
  parseSlots,
  serializeSlots,
  slotLabel,
} from "../paintbrush";
import { getElement } from "../ui/dom";

const PAINT_KEY = "plsfix.paint.v1";
let paintSlots: PaintSlots = emptySlots();

export function loadPaintSlots(): void {
  let raw: string | null = null;
  try {
    raw = localStorage.getItem(PAINT_KEY);
  } catch {
    raw = null;
  }
  paintSlots = parseSlots(raw);
}

export function renderPaintSlots(): void {
  paintSlots.forEach((slot, index) => {
    getElement(`paint-slot-${index + 1}`).textContent = slotLabel(slot);
  });
}

export async function capturePaintSlot(index: number): Promise<string> {
  const slot = await captureSlot(index);
  paintSlots[index - 1] = slot;
  try {
    localStorage.setItem(PAINT_KEY, serializeSlots(paintSlots));
  } catch {
    // Storage can be unavailable in private webviews; slots stay in memory.
  }
  renderPaintSlots();
  return `Slot ${index}: ${slotLabel(slot)}`;
}

export async function applyPaintSlot(index: number): Promise<string> {
  await applySlot(index, paintSlots[index - 1] ?? null);
  return `Painted slot ${index}`;
}
