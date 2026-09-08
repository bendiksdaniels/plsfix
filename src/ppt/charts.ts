// Chart links on a slide: what a payload's chart data would cost to draw,
// whether this host can afford it, the insert that lands it as a tagged group
// of native shapes, the repaint that rebuilds it where the user left it, and
// the fall back to a picture when the source stopped being drawable.
// Invariant: a link drawn as a group refreshes as a group, and only as one.

import { chartSize, layoutChart, type Primitive } from "../chart-shapes";
import type { Box, Size } from "../layout";
import {
  CHART_HOST_SILENT,
  pictureNote,
  type ChartData,
} from "../link/chart-model";
import {
  encodeTag,
  sourceLabel,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
  type Payload,
  type PicturePayload,
} from "../link/model";
import { base64ToBytes, pngSize } from "../link/png";
import { drawGroup, isDrawTimeout, withSyncDeadline } from "./chart-draw";
import type { FoundLink, InsertResult } from "./host";
import {
  CONTENT_WIDTH,
  placeOnSlide,
  selectedSlideId,
  SLIDE,
} from "./placement";
import { hasPowerPointApi, isGrouped } from "./shapes";

export { SHAPES_PER_SYNC } from "./chart-draw";

// Shape groups, text boxes and the formats a chart wears arrived in
// PowerPointApi 1.8; a pie's start and end angle need 1.10.
const CHART_API = "1.8";
const PIE_API = "1.10";
export const CHARTS_NEED_1_8 =
  "as a picture: shape charts need PowerPoint 2504/16.96 or newer";
export const PIES_NEED_1_10 =
  "as a picture: pie shapes need PowerPoint 2601/16.105 or newer";

// How many shapes a host draws before the cost of an add outgrows the value
// of drawing at all: on the web an add slows down with every shape already on
// the slide, and a batch of fifty never came back (spike, 30.08).
export const SHAPE_BUDGET_WEB = 30;
export const SHAPE_BUDGET_DESKTOP = 200;

// A picture's pixels are 96 to the inch and a slide's points are 72.
const PNG_TO_POINTS = 0.75;
// The same margin CONTENT_WIDTH keeps, on the other side: a chart taller than
// this is scaled down to it, exactly as fitToSlide scales the picture, or the
// group hangs off the top and the bottom of the slide.
const CONTENT_HEIGHT = SLIDE.height - (SLIDE.width - CONTENT_WIDTH);

export function shapeBudget(): number {
  return Office.context?.platform === Office.PlatformType.OfficeOnline
    ? SHAPE_BUDGET_WEB
    : SHAPE_BUDGET_DESKTOP;
}

export function overBudgetNote(count: number, budget: number): string {
  return `as a picture: ${String(count)} shapes is over this host's budget of ${String(budget)}`;
}

// What a chart payload would become on a slide: the data, the size the picture
// beside it says the chart is, and the primitives that fill a box that size.
export interface ChartPlan {
  data: ChartData;
  size: Size;
  primitives: Primitive[];
  // The picture that travelled beside the data: what goes on the slide when
  // the host turns out not to be able to draw the shapes after all.
  png: string;
}

// chartSize caps the width alone, because how tall a chart may be is the
// slide's business and not the layout's: a chart the sheet made taller than
// the slide is scaled back, aspect kept, before anything is laid out in it.
function onSlide(size: Size): Size {
  const scale = Math.min(1, CONTENT_HEIGHT / size.height);
  return { width: size.width * scale, height: size.height * scale };
}

// Null for every payload with no chart data at all: a table, a plain range,
// and a chart type Excel could not describe, which are pictures and stay so.
export function chartPlan(payload: Payload): ChartPlan | null {
  if (payload.kind !== "picture" || payload.chart === undefined) return null;
  const data = payload.chart;
  const size = onSlide(
    chartSize(
      {
        width: payload.width * PNG_TO_POINTS,
        height: payload.height * PNG_TO_POINTS,
      },
      CONTENT_WIDTH,
    ),
  );
  return {
    data,
    size,
    primitives: layoutChart(data, { left: 0, top: 0, ...size }),
    png: payload.png,
  };
}

// Null means draw it. Anything else is the sentence the pane shows beside the
// picture it inserted instead, and every one of them names the picture first.
export function declineReason(plan: ChartPlan): string | null {
  if (!hasPowerPointApi(CHART_API)) return CHARTS_NEED_1_8;
  if (plan.data.kind === "pie" && !hasPowerPointApi(PIE_API)) {
    return PIES_NEED_1_10;
  }
  const budget = shapeBudget();
  const count = plan.primitives.length;
  return count > budget ? overBudgetNote(count, budget) : null;
}

// The layout at the box the group actually gets: the plan's own primitives
// when the placement kept its size, laid out again when the slide had to
// shrink it or the user left the group at another width.
function primitivesAt(plan: ChartPlan, box: Box): Primitive[] {
  if (box.width === plan.size.width && box.height === plan.size.height) {
    return plan.primitives;
  }
  return layoutChart(plan.data, {
    left: 0,
    top: 0,
    width: box.width,
    height: box.height,
  });
}

// The box a rebuild draws into: the corner and width the user chose, and the
// only geometry a refresh decides for itself - the height the chart's own
// aspect asks for at that width.
function boxAtFound(found: FoundLink, size: Size): Box {
  return {
    left: found.left,
    top: found.top,
    width: found.width,
    height: Math.round(found.width * (size.height / size.width)),
  };
}

function refreshStage(found: FoundLink): string {
  return `refresh ${sourceLabel(found.tag.src, found.tag.kind)}`;
}

// Reinserting a group would drop it on the slide, not back into the group the
// user built, so a chart inside one is left alone and the row says why.
function requireUngrouped(found: FoundLink, stage: string): void {
  if (isGrouped(found)) {
    throw new Error(`${stage}: ungroup the chart before it can update`);
  }
}

// The linked picture as one shape: the rectangle every fallback lands, image
// filled and carrying both tags, so a chart that could not be drawn is still
// a link the deck updates.
function addPicture(
  shapes: PowerPoint.ShapeCollection,
  box: Box,
  label: string,
  png: string,
  identity: { tag: LinkTag; token: string },
): PowerPoint.Shape {
  const shape = shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.rectangle,
    box,
  );
  shape.name = `pls,fix link ${label}`;
  shape.lineFormat.visible = false;
  shape.fill.setImage(png);
  shape.tags.add(TAG_LINK, encodeTag(identity.tag));
  shape.tags.add(TAG_KEY, identity.token);
  return shape;
}

// The host swallowed the draw: the shapes it had taken are already deleted
// (chart-draw.ts), so the picture goes in the space the chart was placed in,
// in a fresh batch and under the same deadline. A host that has stopped
// answering altogether rejects here instead, which still frees the pane.
async function pictureInstead(
  where: { slideId: string; box: Box },
  item: InboxItem,
  png: string,
  tag: LinkTag,
): Promise<string> {
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(where.slideId).shapes;
    const shape = addPicture(shapes, where.box, item.label, png, {
      tag,
      token: item.token,
    });
    shape.load("id");
    await withSyncDeadline(context.sync());
    return shape.id;
  });
}

export async function insertChart(
  stage: string,
  item: InboxItem,
  plan: ChartPlan,
  tag: LinkTag,
): Promise<InsertResult> {
  let where: { slideId: string; box: Box; overlapping: boolean } | undefined;
  try {
    return await PowerPoint.run(async (context) => {
      // The reads that decide where the chart goes are round trips like any
      // other, and a host that swallows one of them would leave the pane
      // waiting for ever: they get the draw's own deadline.
      const slideId = await withSyncDeadline(selectedSlideId(context, stage));
      const placed = await withSyncDeadline(
        placeOnSlide(context, slideId, plan.size),
      );
      where = { slideId, ...placed };
      const shapes = context.presentation.slides.getItem(slideId).shapes;
      const shapeId = await drawGroup(context, shapes, {
        primitives: primitivesAt(plan, placed.box),
        box: placed.box,
        font: plan.data.font,
        name: `pls,fix chart ${item.label}`,
        tag,
        token: item.token,
        slideId,
      });
      return { slideId, shapeId, overlapping: placed.overlapping };
    });
  } catch (error) {
    const placed = where;
    if (!isDrawTimeout(error) || placed === undefined) throw error;
    return {
      slideId: placed.slideId,
      shapeId: await pictureInstead(placed, item, plan.png, tag),
      overlapping: placed.overlapping,
      note: pictureNote(CHART_HOST_SILENT),
    };
  }
}

// The chart drawn again where it sits. The new group is built first and the
// old one deleted in the same sync afterwards, the order tables.ts recreates
// in: a host that refuses the add never gets as far as taking the old group -
// and both its tags - down.
export async function refreshChart(
  found: FoundLink,
  plan: ChartPlan,
  tag: LinkTag,
): Promise<void> {
  const stage = refreshStage(found);
  requireUngrouped(found, stage);
  const box = boxAtFound(found, plan.size);
  const label = sourceLabel(found.tag.src, found.tag.kind);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(found.slideId).shapes;
    const old = shapes.getItem(found.shapeId);
    await drawGroup(context, shapes, {
      primitives: primitivesAt(plan, box),
      box,
      font: plan.data.font,
      name: `pls,fix chart ${label}`,
      tag,
      token: found.token,
      slideId: found.slideId,
      before: () => {
        old.delete();
      },
    });
  });
}

// The modeller turned the source into a chart type no slide can draw, or this
// host lost the API that drew it: the picture goes where the group was, in one
// sync, with the add queued before the delete for the same reason as above.
export async function replaceGroupWithPicture(
  found: FoundLink,
  payload: PicturePayload,
  tag: LinkTag,
): Promise<void> {
  const stage = refreshStage(found);
  requireUngrouped(found, stage);
  const box = boxAtFound(found, pngSize(base64ToBytes(payload.png)));
  const label = sourceLabel(found.tag.src, found.tag.kind);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(found.slideId).shapes;
    const old = shapes.getItem(found.shapeId);
    addPicture(shapes, box, label, payload.png, { tag, token: found.token });
    old.delete();
    await withSyncDeadline(context.sync());
  });
}

// What "Update all" does to a link the deck holds as a group: draw the chart
// again when this host still can, and put the picture there when it cannot.
export async function refreshChartGroup(
  found: FoundLink,
  payload: PicturePayload,
  tag: LinkTag,
): Promise<string | undefined> {
  const plan = chartPlan(payload);
  let silent = false;
  if (plan !== null && declineReason(plan) === null) {
    try {
      await refreshChart(found, plan, tag);
      return undefined;
    } catch (error) {
      // A host that swallowed the redraw keeps the group it already had: the
      // picture replaces it at the same corner and width, so the link is
      // still a link. Anything else is the host's own refusal, unchanged.
      if (!isDrawTimeout(error)) throw error;
      silent = true;
    }
  }
  await replaceGroupWithPicture(found, payload, tag);
  // The reason travels back so "Update all" can say why a group is a picture.
  return silent ? pictureNote(CHART_HOST_SILENT) : undefined;
}
