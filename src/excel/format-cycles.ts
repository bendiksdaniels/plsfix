// The format cycles: fill, font colour, row style, number format, borders and
// the three hygiene ones (indent, alignment, underline). Each one reads its
// current state from the selection's first cell (borders from the first area's
// own edges), steps once through the brand's cycle and writes the next look
// into every area of the selection, then answers with the step it landed on
// (src/cycle-labels.ts names it; a mixed selection reads off the one cell the
// cycle's state comes from, so the receipt always names a real step).
//
// Owns: the Office.js side of the cycles only; which looks follow which lives
// in the pure src/cycles.ts, the labels in src/cycle-labels.ts, and the
// presets in selection.ts.

import { activeArea, cappedAreas, selectedAreas } from "./areas";
import { protectedNote, syncWrite } from "./protection";
import { captureUndoAreas } from "./undo";
import {
  alignLabel,
  BORDER_CYCLE_LABELS,
  FILL_CYCLE_LABELS,
  FONT_CYCLE_LABELS,
  numberCycleLabels,
  ROW_STYLE_LABELS,
  underlineLabel,
} from "../cycle-labels";
import {
  BORDER_EDGE_NAMES,
  type BorderCycleState,
  type BorderEdgeName,
  type BorderReadouts,
  type BorderSpec,
  buildBorderCycle,
  buildFillCycle,
  buildFontCycle,
  buildNumberCycles,
  buildRowStyleCycles,
  type CellStyle,
  CLEAR_FILL,
  matchBorderIndex,
  matchStyleIndex,
  nextAlignment,
  nextIndent,
  nextInCycle,
  nextUnderline,
  type NumberCycleFamily,
  type RowStyleKind,
  type StyleSpec,
} from "../cycles";
import { makeFormatGrid } from "../model";
import { getActiveSettings } from "../settings";

const ROW_STYLE_ROW_CAP = 500;

// Cycle state lives in the cell: every run reads the active cell and steps once.
function readFill(cell: Excel.Range): string {
  const { fill } = cell.format;
  // Excel reports white for unfilled cells, so the pattern decides.
  return fill.pattern === Excel.FillPattern.none
    ? CLEAR_FILL
    : fill.color.toUpperCase();
}

function readCellStyle(cell: Excel.Range): CellStyle {
  const fill = readFill(cell);
  return {
    fill: fill === CLEAR_FILL ? null : fill,
    fontColor: cell.format.font.color.toUpperCase(),
    bold: cell.format.font.bold,
  };
}

function applyBorder(
  format: Excel.RangeFormat,
  index: Excel.BorderIndex,
  spec: BorderSpec | null | undefined,
): void {
  if (spec === undefined) return;
  const border = format.borders.getItem(index);
  if (spec === null) {
    border.style = Excel.BorderLineStyle.none;
    return;
  }
  if (spec.style === "double") {
    border.style = Excel.BorderLineStyle.double;
  } else {
    border.style = Excel.BorderLineStyle.continuous;
    border.weight = Excel.BorderWeight.thin;
  }
  border.color = spec.color;
}

function applyStyleSpec(format: Excel.RangeFormat, spec: StyleSpec): void {
  if (spec.fill === CLEAR_FILL) format.fill.clear();
  else if (spec.fill !== undefined) format.fill.color = spec.fill;
  if (spec.fontColor !== undefined) format.font.color = spec.fontColor;
  if (spec.bold !== undefined) format.font.bold = spec.bold;
  applyBorder(format, Excel.BorderIndex.edgeTop, spec.topBorder);
  applyBorder(format, Excel.BorderIndex.edgeBottom, spec.bottomBorder);
}

export async function applyNumberCycle(
  family: NumberCycleFamily,
): Promise<string> {
  return Excel.run(async (context) => {
    const areas = await cappedAreas(context, "Format cycling");
    const active = activeArea(areas).getCell(0, 0);
    for (const area of areas) area.load("rowCount,columnCount");
    active.load("numberFormat");
    await context.sync();
    await captureUndoAreas(context, areas);

    const settings = getActiveSettings();
    const cycle = buildNumberCycles(settings)[family];
    const current = active.numberFormat[0]?.[0];
    const next = nextInCycle(typeof current === "string" ? current : "", cycle);
    const label = numberCycleLabels(settings)[family][cycle.indexOf(next)];

    for (const area of areas) {
      area.numberFormat = makeFormatGrid(area.rowCount, area.columnCount, next);
    }
    await syncWrite(context, "Format cycling", protectedNote, areas.length);
    return `Number format: ${label ?? next}`;
  });
}

export async function applyRowStyleCycle(kind: RowStyleKind): Promise<string> {
  return Excel.run(async (context) => {
    const areas = await selectedAreas(context, "Row styles");
    const active = activeArea(areas).getCell(0, 0);
    for (const area of areas) area.load("rowCount");
    active.load(
      "format/fill/color,format/fill/pattern,format/font/color,format/font/bold",
    );
    await context.sync();

    // Edge borders target the whole range, which would leave interior rows
    // bare in a multi-row selection; row styles are per-row by definition.
    // Refuse absurd heights instead of silently degrading to edge borders -
    // before the capture, so a refusal that writes nothing neither spends an
    // undo slot nor drops the stack for being over the capture's cell cap.
    const rows = areas.reduce((total, area) => total + area.rowCount, 0);
    if (rows > ROW_STYLE_ROW_CAP) {
      throw new Error(
        `Row styles support up to ${ROW_STYLE_ROW_CAP} rows at once.`,
      );
    }

    const variants = buildRowStyleCycles(getActiveSettings())[kind];
    const index = matchStyleIndex(readCellStyle(active), variants);
    const nextIndex = (index + 1) % variants.length;
    const next = variants[nextIndex];
    await captureUndoAreas(context, areas);

    if (next) for (const area of areas) paintRowStyle(area, next);
    await syncWrite(context, "Row styles", protectedNote, areas.length);
    return `Row style: ${ROW_STYLE_LABELS[kind][nextIndex]}`;
  });
}

function paintRowStyle(area: Excel.Range, next: StyleSpec): void {
  if (area.rowCount > 1) {
    for (let row = 0; row < area.rowCount; row += 1) {
      applyStyleSpec(area.getRow(row).format, next);
    }
    return;
  }
  applyStyleSpec(area.format, next);
}

export async function applyFillCycle(): Promise<string> {
  return Excel.run(async (context) => {
    const areas = await selectedAreas(context, "Fill cycling");
    const active = activeArea(areas).getCell(0, 0);
    active.load("format/fill/color,format/fill/pattern");
    await context.sync();

    const cycle = buildFillCycle(getActiveSettings());
    const next = nextInCycle(readFill(active), cycle);
    await captureUndoAreas(context, areas);

    for (const area of areas) {
      if (next === CLEAR_FILL) area.format.fill.clear();
      else area.format.fill.color = next;
    }

    await syncWrite(context, "Fill cycling", protectedNote, areas.length);
    const label = FILL_CYCLE_LABELS[cycle.indexOf(next)] ?? next;
    return `Fill: ${label}`;
  });
}

export async function applyFontColorCycle(): Promise<string> {
  return Excel.run(async (context) => {
    const areas = await selectedAreas(context, "Font colour cycling");
    const active = activeArea(areas).getCell(0, 0);
    active.load("format/font/color");
    await context.sync();

    const cycle = buildFontCycle(getActiveSettings());
    const next = nextInCycle(active.format.font.color.toUpperCase(), cycle);
    await captureUndoAreas(context, areas);

    for (const area of areas) area.format.font.color = next;
    await syncWrite(
      context,
      "Font colour cycling",
      protectedNote,
      areas.length,
    );
    const label = FONT_CYCLE_LABELS[cycle.indexOf(next)] ?? next;
    return `Font colour: ${label}`;
  });
}

// Borders belong to the selection, not to the active cell: the state is read
// from the first area's own edges and written back to every area's.
interface EdgeHandle {
  edge: BorderEdgeName;
  border: Excel.RangeBorder;
}

function borderIndexes(): Record<BorderEdgeName, Excel.BorderIndex> {
  return {
    top: Excel.BorderIndex.edgeTop,
    bottom: Excel.BorderIndex.edgeBottom,
    left: Excel.BorderIndex.edgeLeft,
    right: Excel.BorderIndex.edgeRight,
    insideHorizontal: Excel.BorderIndex.insideHorizontal,
    insideVertical: Excel.BorderIndex.insideVertical,
  };
}

// An inside line only exists where there is something between: asking a single
// cell for one is an error in Excel, so those edges are left out entirely.
function borderHandles(
  range: Excel.Range,
  rowCount: number,
  columnCount: number,
): EdgeHandle[] {
  const indexes = borderIndexes();
  return BORDER_EDGE_NAMES.filter((edge) => {
    if (edge === "insideHorizontal") return rowCount > 1;
    if (edge === "insideVertical") return columnCount > 1;
    return true;
  }).map((edge) => ({
    edge,
    border: range.format.borders.getItem(indexes[edge]),
  }));
}

function readEdges(handles: EdgeHandle[]): BorderReadouts {
  const readouts: BorderReadouts = {};
  for (const { edge, border } of handles) {
    // A range whose cells disagree reports nothing for that edge.
    readouts[edge] = {
      style: border.style ?? "",
      weight: border.weight ?? "",
      color: border.color ?? "",
    };
  }
  return readouts;
}

function writeEdges(handles: EdgeHandle[], state: BorderCycleState): void {
  for (const { edge, border } of handles) {
    const line = state.find((entry) => entry.edge === edge);
    // Edges this look does not draw are cleared, so stepping never leaves a
    // line from the previous look behind.
    if (!line) {
      border.style = Excel.BorderLineStyle.none;
      continue;
    }
    if (line.style === "double") {
      // Excel draws a double rule at its own weight; setting one is refused.
      border.style = Excel.BorderLineStyle.double;
    } else {
      border.style = Excel.BorderLineStyle.continuous;
      border.weight =
        line.weight === "medium"
          ? Excel.BorderWeight.medium
          : Excel.BorderWeight.thin;
    }
    border.color = line.color;
  }
}

// The hygiene cycles read one scalar off the active cell and write one scalar
// into every area, exactly as the font colour cycle does, so the three share
// their whole shape: what to load, how to step it, where to put it back.
interface HygieneCycle<T> {
  stage: string;
  /** The load path of the one property the position is read from. */
  property: string;
  read: (format: Excel.RangeFormat) => T;
  step: (current: T) => T;
  write: (format: Excel.RangeFormat, next: T) => void;
}

async function applyHygieneCycle<T>(cycle: HygieneCycle<T>): Promise<T> {
  return Excel.run(async (context) => {
    const areas = await selectedAreas(context, cycle.stage);
    const active = activeArea(areas).getCell(0, 0);
    active.load(cycle.property);
    await context.sync();

    const next = cycle.step(cycle.read(active.format));
    await captureUndoAreas(context, areas);

    for (const area of areas) cycle.write(area.format, next);
    await syncWrite(context, cycle.stage, protectedNote, areas.length);
    return next;
  });
}

export async function applyIndentCycle(): Promise<string> {
  const next = await applyHygieneCycle({
    stage: "Indent cycling",
    property: "format/indentLevel",
    read: (format) => format.indentLevel,
    step: nextIndent,
    write: (format, next) => {
      format.indentLevel = next;
    },
  });
  return `Indent: ${next}`;
}

export async function applyAlignmentCycle(): Promise<string> {
  const next = await applyHygieneCycle({
    stage: "Alignment cycling",
    property: "format/horizontalAlignment",
    read: (format) => format.horizontalAlignment as string,
    step: nextAlignment,
    write: (format, next) => {
      format.horizontalAlignment = next as Excel.HorizontalAlignment;
    },
  });
  return `Aligned: ${alignLabel(next)}`;
}

export async function applyUnderlineCycle(): Promise<string> {
  const next = await applyHygieneCycle({
    stage: "Underline cycling",
    property: "format/font/underline",
    read: (format) => format.font.underline as string,
    step: nextUnderline,
    write: (format, next) => {
      format.font.underline = next as Excel.RangeUnderlineStyle;
    },
  });
  return `Underline: ${underlineLabel(next)}`;
}

export async function applyBorderCycle(): Promise<string> {
  return Excel.run(async (context) => {
    const areas = await cappedAreas(context, "Border cycling");
    for (const area of areas) area.load("rowCount,columnCount");
    await context.sync();

    const handles = areas.map((area) =>
      borderHandles(area, area.rowCount, area.columnCount),
    );
    for (const edges of handles) {
      for (const { border } of edges) border.load("style,color,weight");
    }
    await context.sync();

    const states = buildBorderCycle(getActiveSettings());
    const index = matchBorderIndex(readEdges(handles[0] ?? []), states);
    const nextIndex = (index + 1) % states.length;
    const next = states[nextIndex];
    await captureUndoAreas(context, areas);

    if (next) for (const edges of handles) writeEdges(edges, next);
    await syncWrite(context, "Border cycling", protectedNote, areas.length);
    return `Borders: ${BORDER_CYCLE_LABELS[nextIndex]}`;
  });
}
