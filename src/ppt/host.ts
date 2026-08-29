// The only PowerPoint Office.js code: scan the deck for shapes carrying the
// link tags, insert a linked picture, repaint one or a whole batch of them in
// place (or reinsert on hosts below PowerPointApi 1.8), re-point one at another
// link by rewriting its tags, break a link by dropping them, and read or set
// which slide is active. Identity is always the PLSFIX_LINK tag - never a shape
// id, name or position. Every flow here is counted in round trips: one sync
// per batch, never one per shape.

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
import {
  expandGroups,
  GROUP_API,
  hasPowerPointApi,
  isGrouped,
  shapeAt,
  SHAPE_PROPERTIES,
  type PlacedShape,
  type ShapePath,
} from "./shapes";

export interface FoundLink extends ShapePath {
  slideIndex: number;
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

const TAG_PROPERTIES = "items/key,items/value";

// fill.setImage arrived in PowerPointApi 1.8. An older host repaints by
// deleting the shape and inserting the picture again at the same box.
export function supportsInPlaceRefresh(): boolean {
  return hasPowerPointApi(GROUP_API);
}

interface TaggedShape extends PlacedShape {
  tags: PowerPoint.TagCollection;
}

// Three round trips, and office.js allows no fewer: the slides, then every
// slide's shapes, then every shape's tags - the only place identity is read.
// Groups add one sync per nesting level, between the second and the third.
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
    const placed = sets.flatMap((set) =>
      set.shapes.items.map((shape): PlacedShape => ({
        slideId: set.slideId,
        slideIndex: set.slideIndex,
        shape,
        groupPath: [],
      })),
    );
    const tagged = (await expandGroups(context, placed)).map(
      (entry): TaggedShape => {
        const tags = entry.shape.tags;
        tags.load(TAG_PROPERTIES);
        return { ...entry, tags };
      },
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
    groupPath: entry.groupPath.length > 0 ? entry.groupPath : undefined,
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

async function readSelectedSlideId(
  context: PowerPoint.RequestContext,
): Promise<string | null> {
  const selected = context.presentation.getSelectedSlides();
  selected.load("items/id");
  await context.sync();
  return selected.items[0]?.id ?? null;
}

async function selectedSlideId(
  context: PowerPoint.RequestContext,
  stage: string,
): Promise<string> {
  const id = await readSelectedSlideId(context);
  if (!id) throw new Error(`${stage}: select a slide first.`);
  return id;
}

// What "Update this slide" acts on: PowerPoint's own selection, not a tick in
// the pane - the pane cannot see the selection any other way. Null when
// nothing is selected, so the caller can say so instead of guessing a slide.
export async function activeSlideId(): Promise<string | null> {
  return PowerPoint.run((context) => readSelectedSlideId(context));
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
    shape.name = `pls,fix link ${item.label}`;
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

export interface RefreshRequest {
  found: FoundLink;
  payload: Payload;
  rev: number;
}

// One link repainted: the batch of one on a host with fill.setImage, and the
// reinsertion fallback below it.
export async function refreshLink(
  found: FoundLink,
  payload: Payload,
  rev: number,
): Promise<void> {
  if (supportsInPlaceRefresh()) {
    await refreshLinks([{ found, payload, rev }]);
    return;
  }
  const stage = `refresh ${sourceLabel(found.tag.src, found.tag.kind)}`;
  // Reinsertion drops the picture on the slide, not back into its group, so
  // a grouped link is left alone and the row says why.
  if (isGrouped(found)) {
    throw new Error(
      `${stage}: grouped pictures need PowerPoint 2504/16.96 or newer`,
    );
  }
  const size = pngSize(base64ToBytes(payload.png));
  const tag = tagFor(found.tag, payload, rev);
  await reinsertLink(
    stage,
    found,
    payload.png,
    tag,
    refreshedHeight(found, size),
  );
}

// Every in-place repaint of an "Update all" in one round trip: the pictures,
// the tags and the heights are queued for the whole batch and sent with a
// single sync, because a sync per shape is what makes a sixty-link deck crawl.
// False when the host has no fill.setImage, so the caller repaints row by row
// through the reinsertion above. A host that refuses one shape rejects the
// whole batch; retrying those rows one at a time is safe, because a repaint
// writes the same picture, tag and height however often it runs.
export async function refreshLinks(batch: RefreshRequest[]): Promise<boolean> {
  if (!supportsInPlaceRefresh()) return false;
  if (batch.length === 0) return true;
  await PowerPoint.run(async (context) => {
    for (const entry of batch) queueRefresh(context, entry);
    await context.sync();
  });
  return true;
}

// One picture repainted where it sits. Nothing here reads a shape property, so
// no entry in the batch needs a load: the geometry it compares against is the
// one the scan already read.
function queueRefresh(
  context: PowerPoint.RequestContext,
  { found, payload, rev }: RefreshRequest,
): void {
  const height = refreshedHeight(found, pngSize(base64ToBytes(payload.png)));
  const shape = shapeAt(context, found);
  shape.fill.setImage(payload.png);
  if (height !== found.height) shape.height = height;
  shape.tags.add(TAG_LINK, encodeTag(tagFor(found.tag, payload, rev)));
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

// Re-pointing a link at another export: both tags rewritten in one sync, and
// nothing else - no picture, no geometry, no z-order. The shape keeps
// everything the user gave it and only changes what it tracks.
export async function retagLink(
  found: FoundLink,
  tag: LinkTag,
  token: string,
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const tags = shapeAt(context, found).tags;
    tags.add(TAG_LINK, encodeTag(tag));
    tags.add(TAG_KEY, token);
    await context.sync();
  });
}

// Breaking a link leaves the picture exactly where it is; only the identity
// and the key go, so nothing in the deck ever refreshes it again.
export async function breakLink(found: FoundLink): Promise<void> {
  await PowerPoint.run(async (context) => {
    const tags = shapeAt(context, found).tags;
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
