// The format cycles: fill, font colour, row style, number format and borders.
// Each one reads its current state from the selection's first cell (borders
// from the first area's own edges), steps once through the brand's cycle and
// writes the next look into every area of the selection.
//
// Owns: the Office.js side of the cycles only; which looks follow which lives
// in the pure src/cycles.ts, and the presets in selection.ts.

import { activeArea, cappedAreas, selectedAreas } from "./areas";
import { captureUndoAreas } from "./undo";
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
  nextInCycle,
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
): Promise<void> {
  await Excel.run(async (context) => {
    const areas = await cappedAreas(context, "Format cycling");
    const active = activeArea(areas).getCell(0, 0);
    for (const area of areas) area.load("rowCount,columnCount");
    active.load("numberFormat");
    await context.sync();
    await captureUndoAreas(context, areas);

    const current = active.numberFormat[0]?.[0];
    const next = nextInCycle(
      typeof current === "string" ? current : "",
      buildNumberCycles(getActiveSettings())[family],
    );

    for (const area of areas) {
      area.numberFormat = makeFormatGrid(area.rowCount, area.columnCount, next);
    }
    await context.sync();
  });
}

export async function applyRowStyleCycle(kind: RowStyleKind): Promise<void> {
  await Excel.run(async (context) => {
    const areas = await selectedAreas(context, "Row styles");
    const active = activeArea(areas).getCell(0, 0);
    for (const area of areas) area.load("rowCount");
    active.load(
      "format/fill/color,format/fill/pattern,format/font/color,format/font/bold",
    );
    await context.sync();

    const variants = buildRowStyleCycles(getActiveSettings())[kind];
    const index = matchStyleIndex(readCellStyle(active), variants);
    const next = variants[(index + 1) % variants.length];
    await captureUndoAreas(context, areas);

    // Edge borders target the whole range, which would leave interior rows
    // bare in a multi-row selection; row styles are per-row by definition.
    // Refuse absurd heights instead of silently degrading to edge borders.
    const rows = areas.reduce((total, area) => total + area.rowCount, 0);
    if (next && rows > ROW_STYLE_ROW_CAP) {
      throw new Error(
        `Row styles support up to ${ROW_STYLE_ROW_CAP} rows at once.`,
      );
    }
    if (next) for (const area of areas) paintRowStyle(area, next);

    await context.sync();
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

export async function applyFillCycle(): Promise<void> {
  await Excel.run(async (context) => {
    const areas = await selectedAreas(context, "Fill cycling");
    const active = activeArea(areas).getCell(0, 0);
    active.load("format/fill/color,format/fill/pattern");
    await context.sync();

    const next = nextInCycle(
      readFill(active),
      buildFillCycle(getActiveSettings()),
    );
    await captureUndoAreas(context, areas);

    for (const area of areas) {
      if (next === CLEAR_FILL) area.format.fill.clear();
      else area.format.fill.color = next;
    }

    await context.sync();
  });
}

export async function applyFontColorCycle(): Promise<void> {
  await Excel.run(async (context) => {
    const areas = await selectedAreas(context, "Font colour cycling");
    const active = activeArea(areas).getCell(0, 0);
    active.load("format/font/color");
    await context.sync();

    const next = nextInCycle(
      active.format.font.color.toUpperCase(),
      buildFontCycle(getActiveSettings()),
    );
    await captureUndoAreas(context, areas);

    for (const area of areas) area.format.font.color = next;
    await context.sync();
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

export async function applyBorderCycle(): Promise<void> {
  await Excel.run(async (context) => {
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
    const next = states[(index + 1) % states.length];
    await captureUndoAreas(context, areas);

    if (next) for (const edges of handles) writeEdges(edges, next);
    await context.sync();
  });
}
