// The linked picture every chart-drawing fallback lands: one tagged
// rectangle, image filled, and the two ways charts.ts reaches it - a fresh
// batch once the group is already gone (a swallowed draw, chart-cleanup.ts
// already ran) or queued inside the same batch as the draw when the slide
// handed back a box below the minimum. Split out of charts.ts at its 400-line
// cap; no logic changed.

import type { Box } from "../layout";
import {
  encodeTag,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
} from "../link/model";
import { withSyncDeadline } from "./chart-draw";

// The linked picture as one shape: the rectangle every fallback lands, image
// filled and carrying both tags, so a chart that could not be drawn is still
// a link the deck updates.
export function addPicture(
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
export async function pictureInstead(
  where: { slideId: string; box: Box },
  item: InboxItem,
  png: string,
  tag: LinkTag,
): Promise<string> {
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(where.slideId).shapes;
    return pictureSynced(context, shapes, where.box, item, png, tag);
  });
}

// The picture added, its id read back and the batch committed under the
// deadline in the caller's run: the timeout fallback and the too-small
// refusal land the same shape.
export async function pictureSynced(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection,
  box: Box,
  item: InboxItem,
  png: string,
  tag: LinkTag,
): Promise<string> {
  const shape = addPicture(shapes, box, item.label, png, {
    tag,
    token: item.token,
  });
  shape.load("id");
  await withSyncDeadline(context.sync());
  return shape.id;
}
