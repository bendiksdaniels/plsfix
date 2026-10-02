// Paintbrush, Macabacus style: three format slots that carry a cell's look
// without the clipboard. Capture reads the active cell into a PaintSlot (plain
// data, src/paintbrush.ts); apply paints one slot over every cell of the
// selection in a single setCellProperties write, behind pls,fix Undo and the
// selection cap. Only formatting travels - values and formulas are never read
// and never written.

import { readFill, requireNoOverlayOwner } from "./fill-store";
import {
  BASE_WHITE,
  selectedSingleRange,
  settableEdge,
  withinCap,
} from "./internal";
import { syncWrite } from "./protection";
import { captureUndo } from "./undo";
import { makeFormatGrid } from "../model";
import {
  PAINT_SLOT_COUNT,
  parseSlots,
  serializeSlots,
  type PaintBorder,
  type PaintEdge,
  type PaintSlot,
  type PaintSlots,
} from "../paintbrush";

// Errors say which flow asked, the way the other multi-area guards do.
const STAGE = "paintbrush";
const CAP_LABEL = "Paintbrush";
const PAINT_SLOTS_SETTING = "PLSFIX_PAINT_SLOTS_V1";

// Slots normally travel with the workbook. Local storage remains a small
// fallback for hosts that block workbook settings, and for an unsaved new
// workbook before Excel has a file to carry them in.
export async function loadWorkbookPaintSlots(): Promise<PaintSlots | null> {
  return Excel.run(async (context) => {
    const setting =
      context.workbook.settings.getItemOrNullObject(PAINT_SLOTS_SETTING);
    setting.load("isNullObject,value");
    await context.sync();
    if (setting.isNullObject || typeof setting.value !== "string") return null;
    return parseSlots(setting.value);
  });
}

// Never throws: a host that refuses to persist workbook settings still leaves
// the caller's in-memory and localStorage copies as the fallback the comment
// above promises, instead of turning every capture into a pane error. The
// boolean says whether the workbook copy actually landed.
export async function saveWorkbookPaintSlots(
  slots: PaintSlots,
): Promise<boolean> {
  try {
    await Excel.run(async (context) => {
      context.workbook.settings.add(PAINT_SLOTS_SETTING, serializeSlots(slots));
      await context.sync();
    });
    return true;
  } catch {
    return false;
  }
}

// The scalars a slot is made of, in one load: the number format sits on the
// range, everything else on its format tree. Borders load separately - each
// getItem hands back its own object.
const SLOT_PROPERTIES = [
  "numberFormat",
  "format/horizontalAlignment",
  "format/font/name",
  "format/font/size",
  "format/font/bold",
  "format/font/italic",
  "format/font/color",
  "format/fill/color",
  "format/fill/pattern",
].join(",");

type EdgeBorders = Record<PaintEdge, Excel.RangeBorder>;

function assertSlotIndex(index: number): void {
  if (!Number.isInteger(index) || index < 1 || index > PAINT_SLOT_COUNT) {
    throw new Error(`${STAGE}: slot ${index} does not exist`);
  }
}

// Read inside the call, never at module scope: office.js defines the enums, and
// the pane also loads in a plain browser where they do not exist yet.
function edgeBorders(cell: Excel.Range): EdgeBorders {
  const { borders } = cell.format;
  return {
    top: borders.getItem(Excel.BorderIndex.edgeTop),
    bottom: borders.getItem(Excel.BorderIndex.edgeBottom),
    left: borders.getItem(Excel.BorderIndex.edgeLeft),
    right: borders.getItem(Excel.BorderIndex.edgeRight),
  };
}

// Excel answers with nothing for an edge it cannot describe in one word. One
// cell always can, but a host that hedges must still leave a slot that survives
// the round trip through storage, so every field falls back to a real value.
function readBorder(border: Excel.RangeBorder): PaintBorder {
  return {
    style: border.style || Excel.BorderLineStyle.none,
    weight: border.weight || Excel.BorderWeight.thin,
    color: (border.color || "#000000").toUpperCase(),
  };
}

// Null is no fill, by the one read-back rule (readFill): the desktop reports
// an unfilled cell as white, the web as "" and a solid one with no pattern.
// Only the two loaded properties are handed over, never the proxy itself.
function slotFill(fill: Excel.RangeFill): string | null {
  const read = readFill({ color: fill.color, pattern: fill.pattern });
  return read ? (read.color ?? BASE_WHITE).toUpperCase() : null;
}

function readSlot(cell: Excel.Range, borders: EdgeBorders): PaintSlot {
  const { fill, font, horizontalAlignment } = cell.format;
  const format = cell.numberFormat[0]?.[0];

  return {
    numberFormat: typeof format === "string" ? format : "General",
    font: {
      name: font.name,
      size: font.size,
      bold: font.bold,
      italic: font.italic,
      color: font.color.toUpperCase(),
    },
    fill: slotFill(fill),
    horizontalAlignment,
    borders: {
      top: readBorder(borders.top),
      bottom: readBorder(borders.bottom),
      left: readBorder(borders.left),
      right: readBorder(borders.right),
    },
  };
}

// The active cell, not the first cell of the selection: a modeller points at
// the look they mean, and Excel keeps that cell inside whatever is selected.
export async function captureSlot(index: number): Promise<PaintSlot> {
  assertSlotIndex(index);

  return Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.load(SLOT_PROPERTIES);
    const borders = edgeBorders(cell);
    for (const border of Object.values(borders)) {
      border.load("style,color,weight");
    }
    await context.sync();

    return readSlot(cell, borders);
  });
}

// A slot with no fill clears the fill the way pls,fix Undo restores one: the
// None pattern alone, since Excel for the web keeps a solid fill of whatever
// colour rides beside it (rig 27.09: white + None painted white).
function fillProperties(fill: string | null): Excel.CellPropertiesFill {
  if (fill === null) return { pattern: Excel.FillPattern.none };
  return {
    color: fill,
    pattern: Excel.FillPattern.solid,
    patternColor: BASE_WHITE,
  };
}

// Through the one edge rule (settableEdge): on the web a None edge travels as
// its style alone, or the weight and colour a slot keeps for it draw a line.
function borderProperties(border: PaintBorder): Excel.CellBorder {
  const edge = settableEdge({
    style: border.style as Excel.BorderLineStyle,
    weight: border.weight as Excel.BorderWeight,
    color: border.color,
  });
  return edge ?? { style: Excel.BorderLineStyle.none };
}

function cellProperties(slot: PaintSlot): Excel.SettableCellProperties {
  return {
    format: {
      fill: fillProperties(slot.fill),
      font: { ...slot.font },
      // Per cell, not per range edge: a format painter puts the source cell's
      // own rules on every cell it touches, interior ones included.
      borders: {
        top: borderProperties(slot.borders.top),
        bottom: borderProperties(slot.borders.bottom),
        left: borderProperties(slot.borders.left),
        right: borderProperties(slot.borders.right),
      },
      horizontalAlignment:
        slot.horizontalAlignment as Excel.HorizontalAlignment,
    },
  };
}

// setCellProperties wants a grid the exact shape of the range; every cell gets
// the same look, so one description is shared rather than rebuilt per cell.
function propertyGrid(
  rows: number,
  columns: number,
  slot: PaintSlot,
): Excel.SettableCellProperties[][] {
  const properties = cellProperties(slot);
  return Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => properties),
  );
}

export async function applySlot(
  index: number,
  slot: PaintSlot | null,
): Promise<void> {
  assertSlotIndex(index);
  if (!slot) throw new Error(`${STAGE}: slot ${index} is empty`);
  // An overlay hands back exactly what it snapshotted: painting a slot over
  // its tint first would capture that tint into pls,fix Undo as if it were
  // the modeller's own fill, and later restoring it would hand the tint back
  // in the overlay's place. The same refusal pinstripes.ts already makes.
  requireNoOverlayOwner(STAGE);

  await Excel.run(async (context) => {
    const selection = await selectedSingleRange(context, STAGE);
    const range = await withinCap(context, selection, CAP_LABEL);
    range.load("rowCount,columnCount");
    await context.sync();
    await captureUndo(context, range);

    const { rowCount, columnCount } = range;
    range.numberFormat = makeFormatGrid(
      rowCount,
      columnCount,
      slot.numberFormat,
    );
    range.setCellProperties(propertyGrid(rowCount, columnCount, slot));
    await syncWrite(context, STAGE);
  });
}
