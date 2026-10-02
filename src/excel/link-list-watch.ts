// Keeps a caller's own view of the registry honest between pushes: today a
// source that is deleted or moved is not caught until the next push or the
// next time something re-reads the workbook. This registers one
// workbook-wide onChanged handler, for the pane's lifetime - unlike
// link-watch.ts's auto-push there is no on/off switch here, since deciding
// whether an edit is worth acting on is the caller's job, not this file's.

import { Debouncer, type Clock } from "../link/debounce";

export const LIST_WATCH_DELAY_MS = 2_000;

// The delay and the clock are injectable so tests do not have to wait; the
// pane passes neither.
export interface ListWatchOptions {
  delayMs?: number;
  clock?: Clock;
}

// One key is enough: every caller wants to know "something changed, once it
// stops", never which sheet.
const ANY_EDIT = "*";

// Calls back at most once every delayMs once the workbook goes quiet after an
// edit. Registration failure - a host below the event's floor, or the sync
// that adds it never landing - is swallowed the same way watchActiveSheet
// treats one: there is nothing for a caller to do about it but keep working
// without the live view.
export function watchWorksheetEdits(
  onSettled: () => void,
  options: ListWatchOptions = {},
): void {
  const debouncer = new Debouncer(
    options.delayMs ?? LIST_WATCH_DELAY_MS,
    () => {
      onSettled();
    },
    options.clock,
  );
  Excel.run(async (context) => {
    context.workbook.worksheets.onChanged.add(async () => {
      debouncer.touch([ANY_EDIT]);
    });
    await context.sync();
  }).catch(() => undefined);
}
