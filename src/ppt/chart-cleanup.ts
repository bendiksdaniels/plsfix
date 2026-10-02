// Best-effort teardown for an insert that failed partway through: the shapes
// an earlier sync already confirmed are already on the host, and
// PowerPoint.run gives no rollback of its own, so this deletes them in a
// fresh run - by the ids earlier syncs confirmed, by the chart's name for
// the ones no sync ever answered for, or (cleanupByToken) by the one tag no
// two links ever share, for the single-sync inserts that never get as far as
// reading an id back at all. Invariant: a cleanup failure never replaces or
// hides the caller's own error.

import { TAG_KEY } from "../link/model";
import {
  REGISTER_POLL_MS,
  REGISTER_POLL_TRIES,
  registerPollWait,
} from "./shape-ready";

// One poll of `ids`: whichever answer present get deleted (a second sync only
// spent when something was actually there), and whichever still answer null -
// not confirmed gone, just not found this round - come back for the caller
// to ask again.
async function deleteWhicheverAppeared(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection,
  ids: readonly string[],
): Promise<string[]> {
  const found = ids.map((id) => ({
    id,
    shape: shapes.getItemOrNullObject(id),
  }));
  found.forEach(({ shape }) => shape.load("isNullObject"));
  await context.sync();
  const present = found.filter(({ shape }) => !shape.isNullObject);
  const stillNull = found
    .filter(({ shape }) => shape.isNullObject)
    .map(({ id }) => id);
  if (present.length > 0) {
    for (const { shape } of present) shape.delete();
    await context.sync();
  }
  return stillNull;
}

// Deletes every id drawGroup had already confirmed before the sync that
// stopped it, so a failed insert or refresh leaves nothing behind. Runs in
// its own PowerPoint.run because the context that failed may not still be
// usable, and swallows whatever goes wrong here: the caller already has the
// error it means to surface, and a half-done cleanup is still better than
// none. With `namePrefix`, a second pass sweeps the slide's top level for
// every shape named for the chart (drawGroup names each add in the batch
// that adds it): a batch the host applied and then refused - a tier of
// sub-groups, a chunk with one bad primitive - left shapes whose ids nobody
// holds, and a sub-group swallows its members' ids too.
export async function cleanupShapes(
  slideId: string,
  ids: readonly string[],
  namePrefix?: string,
): Promise<void> {
  if (ids.length === 0 && namePrefix === undefined) return;
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItem(slideId);
      let stillNull = await deleteWhicheverAppeared(context, slide.shapes, ids);
      // PowerPoint for the web, rig 27.09 (src/ppt/shape-ready.ts): an id a
      // sync already confirmed can still answer null here for a stretch -
      // not gone, just not on the host's own lookup yet. Poll what is still
      // null, together, at the same interval and cap (~3 s), so a genuine
      // orphan is still found once the host catches up; a clean cleanup -
      // every id resolves on the first check - never enters this loop, so it
      // never costs more than the two syncs it always has.
      for (
        let tries = 0;
        stillNull.length > 0 && tries < REGISTER_POLL_TRIES;
        tries += 1
      ) {
        await registerPollWait(REGISTER_POLL_MS);
        stillNull = await deleteWhicheverAppeared(
          context,
          slide.shapes,
          stillNull,
        );
      }
      if (namePrefix !== undefined) {
        await sweepByName(context, slide.shapes, namePrefix);
      }
    });
  } catch {
    // Best effort: the caller already has the error it means to surface.
  }
}

// A picture, a text box or a table's own add-and-tag sync is a single round
// trip: the add, both tags and shape.load("id") all queue in the one batch
// that sync confirms. A real host may apply that batch and still answer the
// sync() call itself with an error - the same Office.js caveat drawGroup's
// own comment already names for a chart's many syncs - which here leaves a
// fully tagged shape whose id was never read back, because the rejected
// sync's own load("id") delivered nothing to catch a cleanup by id. The
// token is set once per LINK (link-anchors.ts), not per push: a second
// insert of a link already on the slide shares its token with the user's own
// earlier, finished shape. `before` is what makes matching on it safe - only
// a shape whose id was NOT already on the slide when this insert started is
// ever a candidate, so an older copy of the same link is never touched.
export async function cleanupByToken(
  slideId: string,
  token: string,
  before: ReadonlySet<string>,
): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(slideId).shapes;
      shapes.load("items/id");
      await context.sync();
      const candidates = shapes.items.filter((shape) => !before.has(shape.id));
      if (candidates.length === 0) return;
      const tagged = candidates.map((shape) => ({ shape, tags: shape.tags }));
      tagged.forEach(({ tags }) => tags.load("items/key,items/value"));
      await context.sync();
      const strays = tagged.filter(({ tags }) =>
        tags.items.some((tag) => tag.key === TAG_KEY && tag.value === token),
      );
      if (strays.length === 0) return;
      for (const { shape } of strays) shape.delete();
      await context.sync();
    });
  } catch {
    // Best effort: the caller already has the error it means to surface.
  }
}

// The slide's own list is the one place a shape with an unknown id can still
// be found. A link's finished group is named for its chart without the
// prefix's separator, so an older link from the same source is never swept.
async function sweepByName(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection,
  prefix: string,
): Promise<void> {
  shapes.load("items/id,items/name");
  await context.sync();
  const strays = shapes.items.filter((shape) => shape.name.startsWith(prefix));
  if (strays.length === 0) return;
  for (const shape of strays) shape.delete();
  await context.sync();
}
