// Best-effort teardown for a chart draw that failed partway through: the
// shapes drawGroup already synced are already on the host, and PowerPoint.run
// gives no rollback of its own, so this deletes them in a fresh run - by the
// ids earlier syncs confirmed, then by the chart's name for the ones no sync
// ever answered for. Invariant: a cleanup failure never replaces or hides the
// caller's own error.

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
      const found = ids.map((id) => slide.shapes.getItemOrNullObject(id));
      found.forEach((shape) => {
        shape.load("isNullObject");
      });
      await context.sync();
      for (const shape of found) {
        if (!shape.isNullObject) shape.delete();
      }
      await context.sync();
      if (namePrefix !== undefined) {
        await sweepByName(context, slide.shapes, namePrefix);
      }
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
