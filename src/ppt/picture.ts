// Inserting a picture through the Office selection API, the only route on a
// host without fill.setImage (below PowerPointApi 1.8). Owns that call and the
// hunt for the shape it produced. Invariant: the picture lands on the slide it
// is told to, at the box it is given, and the new shape is the last one there.

import type { Box } from "../link/status";
import { withSyncDeadline } from "./chart-draw";

// The slide must be active for setSelectedDataAsync, and the new picture is
// the last shape on it.
export function insertPictureBySelection(
  stage: string,
  slideId: string,
  png: string,
  box: Box,
): Promise<string> {
  return PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    await withSyncDeadline(context.sync(), "inserting the picture");
    await setSelectedPicture(stage, png, box);
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    shapes.load("items/id");
    // The picture already landed through setSelectedDataAsync above; this
    // sync only reads back its id. A host that swallows it never confirms
    // one, so there is nothing here for a cleanup to delete.
    await withSyncDeadline(context.sync(), "inserting the picture");
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
