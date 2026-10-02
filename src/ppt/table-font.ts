// The font size a native table should be created with. Pure: no PowerPoint
// host, no I/O. Owns the modal-size rule tables.ts's addTable and writeCell
// share, split out because tables.ts sits at the file-length cap.

import type { TablePayload } from "../link/model";

/**
 * The size most of a payload's cells carry, or undefined when none of them
 * name one. A tie keeps the smaller size, so a table split evenly between two
 * sizes never rounds up to the larger one's row height.
 */
export function uniformSize(payload: TablePayload): number | undefined {
  const counts = new Map<number, number>();
  for (const row of payload.cells) {
    for (const cell of row) {
      if (cell.z === undefined) continue;
      counts.set(cell.z, (counts.get(cell.z) ?? 0) + 1);
    }
  }
  let modal: number | undefined;
  let best = 0;
  for (const [size, count] of counts) {
    if (
      count > best ||
      (count === best && (modal === undefined || size < modal))
    ) {
      modal = size;
      best = count;
    }
  }
  return modal;
}
