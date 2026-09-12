// The one answer for a shape a row still names and the deck no longer holds
// where it held it: deleted, ungrouped, or dragged into a group since the list
// was drawn. Owns that sentence and the office.js code behind it.
// Invariant: only ItemNotFound is reworded; every other host error travels on
// exactly as the host raised it.

// PowerPoint's own code for "no such slide, shape or tag", on the error object
// and on the OfficeExtension.Error a sync rejects with.
const ITEM_NOT_FOUND = "ItemNotFound";

export function isMissingShape(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    (error as { code?: unknown }).code === ITEM_NOT_FOUND
  );
}

// The pane always refreshes the list after an action, so the next press finds
// the object where the user actually left it; `stage` already names the link.
export function missingShapeError(stage: string, error: unknown): unknown {
  return isMissingShape(error)
    ? new Error(`${stage}: that object is no longer where the list had it.`)
    : error;
}
