// Paintbrush slots as plain data: what a captured cell format is, how three of
// them survive as JSON, and the one line the pane shows for a slot. No Office.js
// and no storage here - src/excel/paintbrush.ts reads and writes Excel, main.ts
// owns the localStorage key.

export const PAINT_SLOT_COUNT = 3;

export const PAINT_EDGES = ["top", "bottom", "left", "right"] as const;
export type PaintEdge = (typeof PAINT_EDGES)[number];

export interface PaintBorder {
  style: string;
  weight: string;
  color: string;
}

export interface PaintFont {
  name: string;
  size: number;
  bold: boolean;
  italic: boolean;
  color: string;
}

export interface PaintSlot {
  numberFormat: string;
  font: PaintFont;
  // Null is an unfilled cell: Excel reports white with pattern None for one, so
  // the colour alone could not tell the two apart.
  fill: string | null;
  horizontalAlignment: string;
  borders: Record<PaintEdge, PaintBorder>;
}

// Always PAINT_SLOT_COUNT long; a null entry is a slot nobody has captured yet.
export type PaintSlots = (PaintSlot | null)[];

export function emptySlots(): PaintSlots {
  return Array.from({ length: PAINT_SLOT_COUNT }, () => null);
}

export function serializeSlots(slots: PaintSlots): string {
  return JSON.stringify(slots.slice(0, PAINT_SLOT_COUNT));
}

// Storage is shared with whatever else the webview kept: anything that is not
// a slot we wrote reads as an empty slot rather than being painted onto cells.
export function parseSlots(raw: string | null): PaintSlots {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw ?? "");
  } catch {
    return emptySlots();
  }
  if (!Array.isArray(parsed)) return emptySlots();

  const slots = emptySlots();
  for (let index = 0; index < PAINT_SLOT_COUNT; index += 1) {
    slots[index] = readSlot(parsed[index]);
  }
  return slots;
}

// What the pane shows under a slot: the two fields a modeller recognises it by.
export function slotLabel(slot: PaintSlot | null): string {
  if (!slot) return "empty";
  return `${slot.numberFormat} · ${slot.font.name}`;
}

function fields(raw: unknown): Record<string, unknown> | null {
  if (typeof raw !== "object" || raw === null) return null;
  return raw as Record<string, unknown>;
}

function readSlot(raw: unknown): PaintSlot | null {
  const source = fields(raw);
  if (!source) return null;

  const font = readFont(source.font);
  const borders = readBorders(source.borders);
  const { numberFormat, horizontalAlignment, fill } = source;
  if (!font || !borders) return null;
  if (typeof numberFormat !== "string") return null;
  if (typeof horizontalAlignment !== "string") return null;
  if (fill !== null && typeof fill !== "string") return null;

  return { numberFormat, font, fill, horizontalAlignment, borders };
}

function readFont(raw: unknown): PaintFont | null {
  const source = fields(raw);
  if (!source) return null;

  const { name, size, bold, italic, color } = source;
  if (typeof name !== "string" || !name) return null;
  if (typeof size !== "number" || !Number.isFinite(size) || size <= 0) {
    return null;
  }
  if (typeof bold !== "boolean" || typeof italic !== "boolean") return null;
  if (typeof color !== "string" || !color) return null;

  return { name, size, bold, italic, color };
}

// All four edges or nothing: a half-read border set would paint one rule and
// silently leave the other three as they were.
function readBorders(raw: unknown): Record<PaintEdge, PaintBorder> | null {
  const source = fields(raw);
  if (!source) return null;

  const top = readBorder(source.top);
  const bottom = readBorder(source.bottom);
  const left = readBorder(source.left);
  const right = readBorder(source.right);
  if (!top || !bottom || !left || !right) return null;

  return { top, bottom, left, right };
}

function readBorder(raw: unknown): PaintBorder | null {
  const source = fields(raw);
  if (!source) return null;

  const { style, weight, color } = source;
  if (typeof style !== "string" || !style) return null;
  if (typeof weight !== "string" || !weight) return null;
  if (typeof color !== "string" || !color) return null;

  return { style, weight, color };
}
