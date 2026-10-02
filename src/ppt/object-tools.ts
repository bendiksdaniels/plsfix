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
import { withSyncDeadline } from "./chart-draw";
import { hasPowerPointApi } from "./shapes";

const GEOMETRY =
  "items/id,items/type,items/left,items/top,items/width,items/height";

// The whole selection surface (getSelectedShapes, setSelectedShapes,
// getParentSlide) is PowerPointApi 1.5; the fill/line members below it are
// 1.4, so gating every entry point at 1.5 covers both.
const OBJECT_TOOLS_API = "1.5";

// Select similar compares the slide's own shapes, which never include one the
// user clicked into a group to reach.
const SIMILAR_NEEDS_SLIDE =
  "Select similar needs an object on the slide, not one inside a group.";

// Exported so placement.ts's "selected shape" target reads the selection
// behind the exact same gate and message, rather than a second copy of it.
export function requireObjectToolsApi(): void {
  if (!hasPowerPointApi(OBJECT_TOOLS_API)) {
    throw new Error("Object tools need PowerPoint 2021 or Microsoft 365.");
  }
}

interface CapturedStyle {
  fill: { type: string; color: string; transparency: number | null };
  line: {
    visible: boolean;
    // Like weight/dashStyle/style below, the host answers null for a line
    // that was never given its own explicit color (still inheriting the
    // theme's), independently of whether the line is visible.
    color: string | null;
    transparency: number | null;
    weight: number | null;
    dashStyle: string | null;
    style: string | null;
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
  requireObjectToolsApi();
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load(GEOMETRY);
    await withSyncDeadline(context.sync(), "reading the selection");
    requireCount(selected.items.length, minimum, message, maximum);
    const moves = make(boxesOf(selected.items));
    writeMoves(selected.items, moves);
    await withSyncDeadline(context.sync(), "moving the shapes");
    return selected.items.length;
  });
}

// The toast names the mode in the pane's own spelling: pptpane.html's select
// labels "center" as "Centre" (object-tools-heading > align-objects), so the
// word the button chose has to read the same way the option that picked it did.
function alignModeLabel(mode: AlignMode): string {
  return mode === "center" ? "centre" : mode;
}

export async function alignSelected(mode: AlignMode): Promise<string> {
  const count = await transform(
    2,
    "Select at least two objects to align.",
    (boxes) => alignBoxes(boxes, mode),
  );
  return `Aligned ${String(count)} objects ${alignModeLabel(mode)}.`;
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
  requireObjectToolsApi();
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load(GEOMETRY);
    await withSyncDeadline(context.sync(), "reading the selection");
    requireCount(selected.items.length, 1, "Select one reference object.", 1);
    const source = selected.items[0]!;
    const slide = source.getParentSlide();
    const shapes = slide.shapes;
    shapes.load(GEOMETRY);
    await withSyncDeadline(context.sync(), "reading the slide's shapes");
    const ids = shapes.items
      .filter((shape) => sameKindAndSize(source, shape))
      .map((shape) => shape.id);
    // The source matches itself, so an empty list means the slide's own
    // collection never saw it: the user clicked into a group. Selecting
    // nothing there would clear their selection and still report a match.
    if (!ids.includes(source.id)) throw new Error(SIMILAR_NEEDS_SLIDE);
    slide.setSelectedShapes(ids);
    await withSyncDeadline(context.sync(), "selecting the similar shapes");
    return `Selected ${String(ids.length)} similar object${ids.length === 1 ? "" : "s"}.`;
  });
}

export async function captureObjectStyle(): Promise<string> {
  requireObjectToolsApi();
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id");
    await withSyncDeadline(context.sync(), "reading the selection");
    requireCount(selected.items.length, 1, "Select one object to capture.", 1);
    const shape = selected.items[0]!;
    shape.fill.load("type,foregroundColor,transparency");
    shape.lineFormat.load("visible,color,transparency,weight,dashStyle,style");
    await withSyncDeadline(context.sync(), "reading the style");
    // The whitelist refuses every other ShapeFillType uniformly - Gradient,
    // Pattern, PictureAndTexture, SlideBackground - so a picture and a
    // gradient meet the same refusal, not two different code paths.
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
        // The host answers null, not the string "null", once the line isn't
        // visible: coercing it through String() here is what used to turn a
        // captured null into a bogus dash/line style applyObjectStyle later
        // wrote back as if it were real.
        dashStyle: shape.lineFormat.dashStyle,
        style: shape.lineFormat.style,
      },
    };
    return "Object style captured.";
  });
}

// A cleared fill has no transparency to restore, and an invisible line has
// no colour, transparency, weight, dashStyle or style: the host answers null
// for all of them, so only a visible capture's line is painted, and only the
// fields that came back as real values on the way in.
function paintShape(shape: PowerPoint.Shape, style: CapturedStyle): void {
  if (style.fill.type === "NoFill") {
    shape.fill.clear();
  } else {
    shape.fill.setSolidColor(style.fill.color);
    if (style.fill.transparency !== null) {
      shape.fill.transparency = style.fill.transparency;
    }
  }
  shape.lineFormat.visible = style.line.visible;
  if (!style.line.visible) return;
  // A visible line whose color was never explicitly set (still the theme
  // default) answers null here, same as an unset weight/dashStyle/style
  // below: writing it back unguarded would paint a literal null onto the
  // target's color the way lessons/2026-09-08 already fixed for its siblings.
  if (style.line.color !== null) shape.lineFormat.color = style.line.color;
  if (style.line.transparency !== null) {
    shape.lineFormat.transparency = style.line.transparency;
  }
  if (style.line.weight !== null) shape.lineFormat.weight = style.line.weight;
  if (style.line.dashStyle !== null) {
    shape.lineFormat.dashStyle = style.line
      .dashStyle as PowerPoint.ShapeLineDashStyle;
  }
  if (style.line.style !== null) {
    shape.lineFormat.style = style.line.style as PowerPoint.ShapeLineStyle;
  }
}

export async function applyObjectStyle(): Promise<string> {
  requireObjectToolsApi();
  if (painter === null) throw new Error("Capture an object style first.");
  const style = painter;
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id");
    await withSyncDeadline(context.sync(), "reading the selection");
    requireCount(
      selected.items.length,
      1,
      "Select at least one target object.",
    );
    for (const shape of selected.items) paintShape(shape, style);
    await withSyncDeadline(context.sync(), "painting the style");
    const count = selected.items.length;
    return `Painted ${String(count)} object${count === 1 ? "" : "s"}.`;
  });
}
