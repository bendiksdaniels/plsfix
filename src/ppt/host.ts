// The PowerPoint Office.js code every link flow goes through: scan the deck
// for shapes carrying the link tags, insert a linked picture (tables.ts the
// table, texts.ts the text box, charts.ts the group), repaint one or a batch
// in place (or reinsert below 1.8), re-point one at another link, break one by
// dropping its tags, and read or set the active slide. Identity is always the
// PLSFIX_LINK tag - never a shape id, name or position; every flow is counted
// in round trips: one sync per batch, never one per shape.

import type { Size } from "../layout";
import {
  decodeTag,
  encodeTag,
  sourceLabel,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
  type Payload,
  type PicturePayload,
} from "../link/model";
import { pictureNote } from "../link/chart-model";
import { base64ToBytes, pngSize } from "../link/png";
import { aspectChanged, fitToSlide } from "../link/status";
import {
  chartPlan,
  declineReason,
  insertChart,
  refreshChartGroup,
} from "./charts";
import { insertPictureBySelection } from "./picture";
import {
  placeOnSlide,
  readSelectedSlideId,
  selectedSlideId,
} from "./placement";
import {
  expandGroups,
  GROUP_API,
  GROUP_TYPE,
  hasPowerPointApi,
  isGrouped,
  shapeAt,
  SHAPE_PROPERTIES,
  type PlacedShape,
  type ShapePath,
} from "./shapes";
import { insertTable, refreshTable } from "./tables";
import { insertText, refreshText } from "./texts";

export interface FoundLink extends ShapePath {
  slideIndex: number;
  tag: LinkTag;
  token: string;
  // What the deck holds this link as: "Group" is a native chart, anything
  // else the picture. What it was inserted as it stays, so a refresh reads it.
  type: string;
  left: number;
  top: number;
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
    type: shape.type,
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

// What "Update this slide" acts on: PowerPoint's own selection, not a tick in
// the pane - the pane cannot see the selection any other way. Null when
// nothing is selected, so the caller can say so instead of guessing a slide.
export async function activeSlideId(): Promise<string | null> {
  return PowerPoint.run((context) => readSelectedSlideId(context));
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

export interface InsertResult {
  slideId: string;
  shapeId: string;
  // True when the slide had no room left and the object sits over what is
  // already there: the pane says so with OVERLAP_NOTE.
  overlapping: boolean;
  // Why a chart arrived as a picture; the pane says it after the overlap note.
  note?: string;
}

// What the pane shows when an insert had to cover something.
export const OVERLAP_NOTE =
  "Placed over other objects: no free space on this slide";

export async function insertLink(
  item: InboxItem,
  payload: Payload,
  rev: number,
): Promise<InsertResult> {
  const stage = `insert ${item.label}`;
  const tag = tagFor(item, payload, rev);
  if (payload.kind === "table") {
    return insertTable(stage, item, payload, tag);
  }
  if (payload.kind === "text") {
    return insertText(stage, item, payload, tag);
  }
  // A chart this host can draw lands as shapes; the rest take the picture,
  // with the reason when there is one.
  const plan = chartPlan(payload);
  const note =
    plan === null ? issueNote(payload) : (declineReason(plan) ?? undefined);
  if (plan !== null && note === undefined) {
    return insertChart(stage, item, plan, tag);
  }
  const size = pngSize(base64ToBytes(payload.png));
  const fitted = fitToSlide(size.width, size.height);
  if (!supportsInPlaceRefresh()) {
    const placed = await PowerPoint.run(async (context) => {
      const slideId = await selectedSlideId(context, stage);
      return { slideId, ...(await placeOnSlide(context, slideId, fitted)) };
    });
    const { slideId, box, overlapping } = placed;
    const shapeId = await insertPictureBySelection(
      stage,
      slideId,
      payload.png,
      box,
    );
    await writeTags(slideId, shapeId, tag, item.token);
    return { slideId, shapeId, overlapping, note };
  }
  return PowerPoint.run(async (context) => {
    const slideId = await selectedSlideId(context, stage);
    const { box, overlapping } = await placeOnSlide(context, slideId, fitted);
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
    return { slideId, shapeId: shape.id, overlapping, note };
  });
}

// Why Excel shipped a chart link with no chart data: the picture is certain,
// and the pane says so the way it says a decline of this host's own.
export function issueNote(payload: PicturePayload): string | undefined {
  return payload.chartIssue === undefined
    ? undefined
    : pictureNote(payload.chartIssue);
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

interface PictureRequest extends RefreshRequest {
  payload: PicturePayload;
}

// A chart group repaints shape by shape, so it leaves the batch like a table.
function isPicture(entry: RefreshRequest): entry is PictureRequest {
  return entry.payload.kind === "picture" && entry.found.type !== GROUP_TYPE;
}

// One link repainted: the batch of one on a host with fill.setImage, and the
// reinsertion fallback below it. Answers with a note when the repaint has
// something to say about what it painted - a chart group that had to become a
// picture is the only one that does - and with nothing when it has not.
export async function refreshLink(
  found: FoundLink,
  payload: Payload,
  rev: number,
): Promise<string | undefined> {
  if (payload.kind === "table") {
    await refreshTable(found, payload, tagFor(found.tag, payload, rev));
    return undefined;
  }
  if (payload.kind === "text") {
    await refreshText(found, payload, tagFor(found.tag, payload, rev));
    return undefined;
  }
  if (found.type === GROUP_TYPE) {
    // charts.ts owns the reason; a version of it with none to give simply
    // answers nothing, and the summary then falls back to its own.
    const note: unknown = await refreshChartGroup(
      found,
      payload,
      tagFor(found.tag, payload, rev),
    );
    return typeof note === "string" ? note : undefined;
  }
  if (supportsInPlaceRefresh()) {
    await refreshLinks([{ found, payload, rev }]);
    return undefined;
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
  return undefined;
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
  // A table is written cell by cell, not with one setImage: a batch holding
  // one goes back to the caller, which replays every row on its own.
  if (!batch.every(isPicture)) return false;
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
  { found, payload, rev }: PictureRequest,
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
