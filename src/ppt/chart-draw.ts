// One chart layout turned into native PowerPoint shapes: a primitive per add,
// the adds chunked into syncs of SHAPES_PER_SYNC, a pie's angles written in
// the sync after the one that created it, and one final sync that groups,
// names and tags the lot. Owns every shape-drawing Office.js call.
// Invariant: the group carries both tags and its children carry none.

import type { Line, Primitive, Text } from "../chart-shapes";
import type { Box } from "../layout";
import { encodeTag, TAG_KEY, TAG_LINK, type LinkTag } from "../link/model";
import { cleanupShapes } from "./chart-cleanup";

// The web charges per shape and per round trip, and a batch of more than a
// dozen adds stops coming back at all (spike, 30.08): twelve is what a slide
// already holding other objects still answers in seconds.
export const SHAPES_PER_SYNC = 12;

// A write batch on PowerPoint for the web can be swallowed whole: it neither
// applies nor rejects, the shapes of the chunks before it stay on the slide
// and the pane waits for ever (lessons, 08.09: a 21-shape pie stopped after
// its first chunk of twelve). Every round trip of a draw therefore has a
// deadline. It is deliberately generous - the web measured twelve rectangles
// in 1.7 s on a clean slide and twenty-four in 15 s at 45 shapes - so only a
// host that has genuinely stopped answering ever reaches it.
export const SYNC_TIMEOUT_MS = 60_000;

// What a round trip that never came back rejects with, so the caller can tell
// a host that stopped answering from one that refused the shapes.
export class ChartDrawTimeout extends Error {
  constructor() {
    super("PowerPoint stopped answering while drawing the chart");
    this.name = "ChartDrawTimeout";
  }
}

export function isDrawTimeout(error: unknown): boolean {
  return error instanceof ChartDrawTimeout;
}

// One piece of host work under that deadline: whatever it answers, unless it
// answers nothing at all. The race keeps a handler on the abandoned promise,
// so a batch that rejects long afterwards is still nobody's unhandled error.
export async function withSyncDeadline<T>(work: Promise<T>): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const deadline = new Promise<never>((_resolve, reject) => {
    timer = setTimeout(() => {
      reject(new ChartDrawTimeout());
    }, SYNC_TIMEOUT_MS);
  });
  try {
    return await Promise.race([work, deadline]);
  } finally {
    clearTimeout(timer);
  }
}

// The teardown runs on the host that has just stopped answering, so it gets
// the same deadline and no more: a cleanup that hangs too would hold the
// pane's busy flag open exactly as the draw did.
async function cleanupWithin(slideId: string, ids: string[]): Promise<void> {
  try {
    await withSyncDeadline(cleanupShapes(slideId, ids));
  } catch {
    // Best effort, the same swallow cleanupShapes makes of its own errors.
  }
}

const ALIGNMENT = { l: "Left", c: "Center", r: "Right" } as const;
// A pie's two adjustment points: the angle it starts at and the one it ends
// at, both degrees clockwise from 3 o'clock.
const WEDGE_START = 0;
const WEDGE_END = 1;

// What one chart draw needs to know: the primitives in the layout's own
// coordinates, the box on the slide they are drawn at, the brand font every
// label wears, the group's name, its two tags, and whatever the caller wants
// queued in the grouping sync (a refresh deletes the old group there).
export interface GroupSpec {
  primitives: Primitive[];
  box: Box;
  font: string;
  name: string;
  tag: LinkTag;
  token: string;
  // The slide the group is drawn on, for the cleanup run a rejection sends
  // after it, which cannot reuse the context that failed.
  slideId: string;
  before?: () => void;
}

interface Added {
  shape: PowerPoint.Shape;
  primitive: Primitive;
}

function offset(inner: Box, box: Box): Box {
  return {
    left: box.left + inner.left,
    top: box.top + inner.top,
    width: inner.width,
    height: inner.height,
  };
}

function chunks<T>(items: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let at = 0; at < items.length; at += size) {
    out.push(items.slice(at, at + size));
  }
  return out;
}

function filled(shape: PowerPoint.Shape, color: string): PowerPoint.Shape {
  shape.fill.setSolidColor(color);
  shape.lineFormat.visible = false;
  return shape;
}

// Four property writes and no more: the font, its size, its colour and the
// alignment. The layout sized the box for PowerPoint's own insets, so the
// margins, the wrap and the autosize are left exactly as the host made them.
function addLabel(
  shapes: PowerPoint.ShapeCollection,
  text: Text,
  font: string,
  at: Box,
): PowerPoint.Shape {
  const shape = shapes.addTextBox(text.text, at);
  const range = shape.textFrame.textRange;
  range.font.name = font;
  range.font.size = text.size;
  range.font.color = text.color;
  range.font.bold = text.bold;
  range.paragraphFormat.horizontalAlignment = ALIGNMENT[text.align];
  return shape;
}

// A box is always normalised to a non-negative width and height (see
// chart-shapes-parts.ts lineBetween), so a rising segment - bottom-left to
// top-right - cannot be drawn as a straight connector from it: the host has
// no "flip" on ConnectorType.straight, only a second preset for the other
// diagonal. Everything else keeps the connector it always drew as.
function addLineShape(
  shapes: PowerPoint.ShapeCollection,
  primitive: Line,
  at: Box,
): PowerPoint.Shape {
  if (primitive.rising) {
    const shape = shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.lineInverse,
      at,
    );
    shape.lineFormat.color = primitive.color;
    shape.lineFormat.weight = primitive.weight;
    return shape;
  }
  const shape = shapes.addLine(PowerPoint.ConnectorType.straight, at);
  // The host reads a zero width or height in the add as "not given" and
  // draws the line sloped over its 72 pt default (PowerPoint for the web,
  // 30.08); written after the add, a zero side sticks.
  shape.width = at.width;
  shape.height = at.height;
  shape.lineFormat.color = primitive.color;
  shape.lineFormat.weight = primitive.weight;
  return shape;
}

// A primitive is one shape: a bar, a legend swatch and a stacked segment are
// rectangles, a slice is a Pie (or an Ellipse when it is the whole circle),
// a baseline or a waterfall connector is a line (straight, or the host's
// other diagonal preset for a rising one).
function addPrimitive(
  shapes: PowerPoint.ShapeCollection,
  primitive: Primitive,
  spec: GroupSpec,
): PowerPoint.Shape {
  const at = offset(primitive.box, spec.box);
  switch (primitive.kind) {
    case "text":
      return addLabel(shapes, primitive, spec.font, at);
    case "line":
      return addLineShape(shapes, primitive, at);
    case "wedge":
      return filled(
        shapes.addGeometricShape(PowerPoint.GeometricShapeType.pie, at),
        primitive.color,
      );
    case "ellipse":
      return filled(
        shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, at),
        primitive.color,
      );
    case "rect":
      return filled(
        shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, at),
        primitive.color,
      );
  }
}

// PowerPoint hands out a shape's adjustments only once it has heard of the
// shape, so these writes are queued after the sync that added the wedges and
// travel with the next one - the following chunk's, or the grouping sync.
function shapeWedges(added: Added[]): void {
  for (const { shape, primitive } of added) {
    if (primitive.kind !== "wedge") continue;
    shape.adjustments.set(WEDGE_START, primitive.start);
    shape.adjustments.set(WEDGE_END, primitive.end);
  }
}

// The whole chart in ceil(primitives / SHAPES_PER_SYNC) + 1 round trips: each
// chunk is added and its ids read back in one sync, and the last sync groups
// them, names the group, tags it and runs whatever the caller queued there.
//
// A sync that rejects after the first one has already put shapes on the
// host: PowerPoint.run does not roll those back, so this tracks every id a
// prior sync confirmed and, on any later rejection, deletes them itself
// (chart-cleanup.ts) before rethrowing the rejection unchanged. A rejection
// on the very first sync has nothing recorded yet, so nothing is cleaned.
// A round trip the host swallows without answering reaches the same path
// through its deadline, as a ChartDrawTimeout the caller falls back on.
export async function drawGroup(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection,
  spec: GroupSpec,
): Promise<string> {
  const ids: string[] = [];
  try {
    for (const chunk of chunks(spec.primitives, SHAPES_PER_SYNC)) {
      const added = chunk.map((primitive): Added => {
        const shape = addPrimitive(shapes, primitive, spec);
        shape.name = `${spec.name}: ${primitive.name}`;
        shape.load("id");
        return { shape, primitive };
      });
      await withSyncDeadline(context.sync());
      ids.push(...added.map((one) => one.shape.id));
      shapeWedges(added);
    }
    const group = shapes.addGroup(ids);
    group.name = spec.name;
    group.tags.add(TAG_LINK, encodeTag(spec.tag));
    group.tags.add(TAG_KEY, spec.token);
    group.load("id");
    spec.before?.();
    await withSyncDeadline(context.sync());
    return group.id;
  } catch (err) {
    if (ids.length > 0) await cleanupWithin(spec.slideId, ids);
    throw err;
  }
}
