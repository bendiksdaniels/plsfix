// PowerPoint selection tools: geometry through object-math, style capture and
// apply through the supported fill/line surface. Every action is one load sync
// and one write sync, regardless of the number of selected shapes.

import {
  alignBoxes,
  distributeBoxes,
  matchSize,
  sameKindAndSize,
  swapBoxes,
  type AlignMode,
  type ObjectBox,
  type ObjectMove,
} from "./object-math";

const GEOMETRY =
  "items/id,items/type,items/left,items/top,items/width,items/height";

interface CapturedStyle {
  fill: { type: string; color: string; transparency: number };
  line: {
    visible: boolean;
    color: string;
    transparency: number;
    weight: number;
    dashStyle: string;
    style: string;
  };
}

let painter: CapturedStyle | null = null;

function boxesOf(shapes: PowerPoint.Shape[]): ObjectBox[] {
  return shapes.map((shape) => ({
    id: shape.id,
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  }));
}

function requireCount(
  count: number,
  minimum: number,
  message: string,
  maximum = Number.POSITIVE_INFINITY,
): void {
  if (count < minimum || count > maximum) throw new Error(message);
}

function writeMoves(shapes: PowerPoint.Shape[], moves: ObjectMove[]): void {
  const byId = new Map(shapes.map((shape) => [shape.id, shape]));
  for (const move of moves) {
    const shape = byId.get(move.id);
    if (!shape) continue;
    if (move.left !== undefined) shape.left = move.left;
    if (move.top !== undefined) shape.top = move.top;
    if (move.width !== undefined) shape.width = move.width;
    if (move.height !== undefined) shape.height = move.height;
  }
}

async function transform(
  minimum: number,
  message: string,
  make: (boxes: ObjectBox[]) => ObjectMove[],
  maximum?: number,
): Promise<number> {
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load(GEOMETRY);
    await context.sync();
    requireCount(selected.items.length, minimum, message, maximum);
    const moves = make(boxesOf(selected.items));
    writeMoves(selected.items, moves);
    await context.sync();
    return selected.items.length;
  });
}

export async function alignSelected(mode: AlignMode): Promise<string> {
  const count = await transform(
    2,
    "Select at least two objects to align.",
    (boxes) => alignBoxes(boxes, mode),
  );
  return `Aligned ${String(count)} objects ${mode}.`;
}

export async function distributeSelected(
  axis: "horizontal" | "vertical",
): Promise<string> {
  const count = await transform(
    3,
    "Select at least three objects to distribute.",
    (boxes) => distributeBoxes(boxes, axis),
  );
  return `Distributed ${String(count)} objects ${axis === "horizontal" ? "across" : "down"}.`;
}

export async function matchSelectedSize(): Promise<string> {
  const count = await transform(
    2,
    "Select the reference object first, then at least one target.",
    matchSize,
  );
  return `Matched ${String(count - 1)} object${count === 2 ? "" : "s"} to the reference size.`;
}

export async function swapSelected(): Promise<string> {
  await transform(2, "Select exactly two objects to swap.", swapBoxes, 2);
  return "Swapped two objects.";
}

export async function selectSimilar(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load(GEOMETRY);
    await context.sync();
    requireCount(selected.items.length, 1, "Select one reference object.", 1);
    const source = selected.items[0]!;
    const slide = source.getParentSlide();
    const shapes = slide.shapes;
    shapes.load(GEOMETRY);
    await context.sync();
    const ids = shapes.items
      .filter((shape) => sameKindAndSize(source, shape))
      .map((shape) => shape.id);
    slide.setSelectedShapes(ids);
    await context.sync();
    return `Selected ${String(ids.length)} similar object${ids.length === 1 ? "" : "s"}.`;
  });
}

export async function captureObjectStyle(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id");
    await context.sync();
    requireCount(selected.items.length, 1, "Select one object to capture.", 1);
    const shape = selected.items[0]!;
    shape.fill.load("type,foregroundColor,transparency");
    shape.lineFormat.load("visible,color,transparency,weight,dashStyle,style");
    await context.sync();
    if (!(["Solid", "NoFill"] as string[]).includes(String(shape.fill.type))) {
      throw new Error(
        "Smart Painter currently supports solid or no-fill objects.",
      );
    }
    painter = {
      fill: {
        type: String(shape.fill.type),
        color: shape.fill.foregroundColor,
        transparency: shape.fill.transparency,
      },
      line: {
        visible: shape.lineFormat.visible,
        color: shape.lineFormat.color,
        transparency: shape.lineFormat.transparency,
        weight: shape.lineFormat.weight,
        dashStyle: String(shape.lineFormat.dashStyle),
        style: String(shape.lineFormat.style),
      },
    };
    return "Object style captured.";
  });
}

export async function applyObjectStyle(): Promise<string> {
  if (painter === null) throw new Error("Capture an object style first.");
  const style = painter;
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id");
    await context.sync();
    requireCount(
      selected.items.length,
      1,
      "Select at least one target object.",
    );
    for (const shape of selected.items) {
      if (style.fill.type === "NoFill") shape.fill.clear();
      else shape.fill.setSolidColor(style.fill.color);
      shape.fill.transparency = style.fill.transparency;
      Object.assign(shape.lineFormat, style.line);
    }
    await context.sync();
    const count = selected.items.length;
    return `Painted ${String(count)} object${count === 1 ? "" : "s"}.`;
  });
}
