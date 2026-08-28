// The only PowerPoint Office.js code: scan the deck for shapes carrying the
// link tags, insert a linked picture, repaint one in place (or reinsert it on
// hosts below PowerPointApi 1.8) and break a link by dropping its tags.
// Identity is always the SMT_LINK tag - never a shape id, name or position.

import {
  decodeTag,
  encodeTag,
  sourceLabel,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
  type Payload,
} from "../link/model";
import { base64ToBytes, pngSize } from "../link/png";
import { aspectChanged, fitToSlide, type Box } from "../link/status";

export interface FoundLink {
  slideId: string;
  slideIndex: number;
  shapeId: string;
  tag: LinkTag;
  token: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

interface Size {
  width: number;
  height: number;
}

const SHAPE_PROPERTIES =
  "items/id,items/left,items/top,items/width,items/height";
const TAG_PROPERTIES = "items/key,items/value";

// fill.setImage arrived in PowerPointApi 1.8. An older host repaints by
// deleting the shape and inserting the picture again at the same box; a host
// that reports no requirements at all is the newest one, the web.
export function supportsInPlaceRefresh(): boolean {
  const requirements = Office.context?.requirements;
  return requirements
    ? requirements.isSetSupported("PowerPointApi", "1.8")
    : true;
}

interface TaggedShape {
  slideId: string;
  slideIndex: number;
  shape: PowerPoint.Shape;
  tags: PowerPoint.TagCollection;
}

// Three round trips, and office.js allows no fewer: the slides, then every
// slide's shapes, then every shape's tags - the only place identity is read.
export async function scanLinks(): Promise<FoundLink[]> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const sets = slides.items.map((slide, slideIndex) => {
      const shapes = slide.shapes;
      shapes.load(SHAPE_PROPERTIES);
      return { slideId: slide.id, slideIndex, shapes };
    });
    await context.sync();
    const tagged = sets.flatMap((set) =>
      set.shapes.items.map((shape): TaggedShape => {
        const tags = shape.tags;
        tags.load(TAG_PROPERTIES);
        return {
          slideId: set.slideId,
          slideIndex: set.slideIndex,
          shape,
          tags,
        };
      }),
    );
    await context.sync();
    return tagged
      .map(toFoundLink)
      .filter((link): link is FoundLink => link !== null);
  });
}

// A shape is a link only when it carries both tags: the identity and the key
// that opens its payload. Anything else on the slide is somebody's picture.
function toFoundLink(entry: TaggedShape): FoundLink | null {
  const items = entry.tags.items;
  const tag = decodeTag(items.find((item) => item.key === TAG_LINK)?.value);
  const token = items.find((item) => item.key === TAG_KEY)?.value;
  if (tag === null || token === undefined || token === "") return null;
  const shape = entry.shape;
  return {
    slideId: entry.slideId,
    slideIndex: entry.slideIndex,
    shapeId: shape.id,
    tag,
    token,
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
}

function tagFor(
  link: { id: string; kind: LinkTag["kind"] },
  payload: Payload,
  rev: number,
): LinkTag {
  return {
    v: 1,
    id: link.id,
    kind: link.kind,
    rev,
    src: payload.src,
    pushedAt: payload.pushedAt,
  };
}

async function selectedSlideId(
  context: PowerPoint.RequestContext,
  stage: string,
): Promise<string> {
  const selected = context.presentation.getSelectedSlides();
  selected.load("items/id");
  await context.sync();
  const id = selected.items[0]?.id;
  if (!id) throw new Error(`${stage}: select a slide first.`);
  return id;
}

// Picture inserted through the selection API (hosts without fill.setImage):
// the slide must be active, and the new picture is the last shape on it.
function insertPictureBySelection(
  stage: string,
  slideId: string,
  png: string,
  box: Box,
): Promise<string> {
  return PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    await context.sync();
    await setSelectedPicture(stage, png, box);
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    shapes.load("items/id");
    await context.sync();
    const id = shapes.items.at(-1)?.id;
    if (!id) {
      throw new Error(`${stage}: PowerPoint reported no inserted picture.`);
    }
    return id;
  });
}

function setSelectedPicture(
  stage: string,
  png: string,
  box: Box,
): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      png,
      {
        coercionType: Office.CoercionType.Image,
        imageLeft: box.left,
        imageTop: box.top,
        imageWidth: box.width,
        imageHeight: box.height,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(`${stage}: ${insertFailure(result.error)}`));
      },
    );
  });
}

function insertFailure(error: Office.Error | undefined): string {
  return error?.message ?? "PowerPoint could not insert the picture.";
}

async function writeTags(
  slideId: string,
  shapeId: string,
  tag: LinkTag,
  token: string,
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shape = context.presentation.slides
      .getItem(slideId)
      .shapes.getItem(shapeId);
    shape.tags.add(TAG_LINK, encodeTag(tag));
    shape.tags.add(TAG_KEY, token);
    await context.sync();
  });
}

export async function insertLink(
  item: InboxItem,
  payload: Payload,
  rev: number,
): Promise<{ slideId: string; shapeId: string }> {
  const stage = `insert ${item.label}`;
  const size = pngSize(base64ToBytes(payload.png));
  const box = fitToSlide(size.width, size.height);
  const tag = tagFor(item, payload, rev);
  if (!supportsInPlaceRefresh()) {
    const slideId = await PowerPoint.run((context) =>
      selectedSlideId(context, stage),
    );
    const shapeId = await insertPictureBySelection(
      stage,
      slideId,
      payload.png,
      box,
    );
    await writeTags(slideId, shapeId, tag, item.token);
    return { slideId, shapeId };
  }
  return PowerPoint.run(async (context) => {
    const slideId = await selectedSlideId(context, stage);
    const shape = context.presentation.slides
      .getItem(slideId)
      .shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, box);
    shape.name = `Model Tools link ${item.label}`;
    shape.lineFormat.visible = false;
    shape.fill.setImage(payload.png);
    shape.tags.add(TAG_LINK, encodeTag(tag));
    shape.tags.add(TAG_KEY, item.token);
    shape.load("id");
    await context.sync();
    return { slideId, shapeId: shape.id };
  });
}

// The one geometry a refresh is allowed to write: a picture whose aspect ratio
// moved gets a height under the width the user chose. Left, top and width are
// the user's, and stay the user's.
function refreshedHeight(found: FoundLink, size: Size): number {
  return aspectChanged(found.width, found.height, size.width, size.height)
    ? Math.round(found.width * (size.height / size.width))
    : found.height;
}

export async function refreshLink(
  found: FoundLink,
  payload: Payload,
  rev: number,
): Promise<void> {
  const stage = `refresh ${sourceLabel(found.tag.src, found.tag.kind)}`;
  const size = pngSize(base64ToBytes(payload.png));
  const tag = tagFor(found.tag, payload, rev);
  const height = refreshedHeight(found, size);
  if (!supportsInPlaceRefresh()) {
    await reinsertLink(stage, found, payload.png, tag, height);
    return;
  }
  await PowerPoint.run(async (context) => {
    const shape = context.presentation.slides
      .getItem(found.slideId)
      .shapes.getItem(found.shapeId);
    shape.fill.setImage(payload.png);
    if (height !== found.height) shape.height = height;
    shape.tags.add(TAG_LINK, encodeTag(tag));
    await context.sync();
  });
}

// The fallback repaint, in the only safe order: the new picture lands on the
// same slide at the same box and is tagged first, and the old shape goes last.
// A failure before the delete leaves a duplicate the user can remove; a delete
// first would lose the picture and both tags for good. The new shape is still
// the last one on the slide while the old one is there, so it is found the
// same way as on a first insert.
async function reinsertLink(
  stage: string,
  found: FoundLink,
  png: string,
  tag: LinkTag,
  height: number,
): Promise<void> {
  const box = { left: found.left, top: found.top, width: found.width, height };
  const shapeId = await insertPictureBySelection(
    stage,
    found.slideId,
    png,
    box,
  );
  await writeTags(found.slideId, shapeId, tag, found.token);
  await PowerPoint.run(async (context) => {
    context.presentation.slides
      .getItem(found.slideId)
      .shapes.getItem(found.shapeId)
      .delete();
    await context.sync();
  });
}

// Breaking a link leaves the picture exactly where it is; only the identity
// and the key go, so nothing in the deck ever refreshes it again.
export async function breakLink(found: FoundLink): Promise<void> {
  await PowerPoint.run(async (context) => {
    const tags = context.presentation.slides
      .getItem(found.slideId)
      .shapes.getItem(found.shapeId).tags;
    tags.delete(TAG_LINK);
    tags.delete(TAG_KEY);
    await context.sync();
  });
}

export async function goToSlide(slideId: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    await context.sync();
  });
}
