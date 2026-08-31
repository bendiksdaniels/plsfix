// Best-effort teardown for a chart draw that failed partway through: the
// shapes drawGroup already synced are already on the host, and PowerPoint.run
// gives no rollback of its own, so this deletes them in a fresh run.
// Invariant: a cleanup failure never replaces or hides the caller's own error.

// Deletes every id drawGroup had already confirmed before the sync that
// stopped it, so a failed insert or refresh leaves nothing behind. Runs in
// its own PowerPoint.run because the context that failed may not still be
// usable, and swallows whatever goes wrong here: the caller already has the
// error it means to surface, and a half-done cleanup is still better than
// none.
export async function cleanupShapes(
  slideId: string,
  ids: readonly string[],
): Promise<void> {
  if (ids.length === 0) return;
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
    });
  } catch {
    // Best effort: the caller already has the error it means to surface.
  }
}
