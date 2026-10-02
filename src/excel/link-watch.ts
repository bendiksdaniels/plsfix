// Auto-push on edit: one workbook-wide onChanged handler that notes where an
// edit landed, waits for the typing to stop and re-pushes the links whose
// anchors it touched. Off by default; the flag lives in the workbook, so a
// reopened file re-arms it from the Links tab. Anchors and rendering stay in
// link-anchors.ts and the push itself in links.ts - this file only decides
// which links a change concerns and when.
//
// Three things it deliberately does not do. The keystroke itself touches no
// Office.js: an edit records its address and the debounce window pays for
// resolving anchors once, so a fast typist cannot queue a sync per keystroke.
// Anchor geometry is therefore read at the end of the window, and a structural
// edit (rows or columns inserted or deleted) re-pushes every link on that sheet
// rather than intersecting addresses against anchors that have since moved. And
// a chart is re-pushed whenever its own worksheet changes, since a chart has no
// range to intersect; one plotted from another sheet's cells needs "Push all".

import { Debouncer, type Clock } from "../link/debounce";
import { intersects } from "../link/geometry";
import type { RegistryEntry } from "../link/model";
import type { RelayApi } from "../link/relay";
import { readRegistry, resolveSources } from "./link-anchors";
import { exclusive } from "./link-lock";
import { pushRegistry, type PushSummary } from "./links";
import { parseAddress, splitAreas } from "./shared";

export const AUTOPUSH_SETTING = "PLSFIX_AUTOPUSH";
export const AUTOPUSH_DELAY_MS = 3_000;

const ENABLED = "1";
// The settings surface this add-in uses has no delete, so "off" is written as
// the empty string, which reads back exactly like a workbook that never had the
// flag at all.
const DISABLED = "";
// Compared as a string: the fake host and older Excel builds deliver the value
// without Excel.DataChangeType existing to compare against.
const RANGE_EDITED = "RangeEdited";

export type Notify = (message: string) => void;

// The delay and the clock are injectable so tests do not have to wait three
// seconds; the pane passes neither.
export interface AutoPushOptions {
  delayMs?: number;
  clock?: Clock;
}

interface SheetEdits {
  // Set by a structural change: every link on the sheet is re-pushed, whatever
  // the addresses say.
  whole: boolean;
  areas: string[];
}

interface Session {
  relay: RelayApi;
  notify: Notify;
  debouncer: Debouncer;
}

let session: Session | null = null;
let handler: OfficeExtension.EventHandlerResult<Excel.WorksheetChangedEventArgs> | null =
  null;
// Toggles are serialized so a fast toggle cannot register the handler twice.
// Pushes are serialized too, but on the queue every link flow shares: a window
// falling due while the Links tab is exporting must not read the registry the
// export is about to add to.
let toggles: Promise<void> = Promise.resolve();
// Worksheet id -> what happened on it since the last flush.
const edits = new Map<string, SheetEdits>();

export function autoPushEnabled(): boolean {
  return handler !== null;
}

export function setAutoPush(
  on: boolean,
  relay: RelayApi,
  notify: Notify,
  options: AutoPushOptions = {},
): Promise<void> {
  // Serialized the way autocolor's handler is, so a fast toggle can never
  // register the handler twice.
  const task = toggles.then(() => apply(on, relay, notify, options));
  toggles = task.catch(() => undefined);
  return task;
}

// Boot: the Links tab asks the workbook whether this file had auto-push on and
// arms it again if so. Returns what the checkbox should show.
export async function restoreAutoPush(
  relay: RelayApi,
  notify: Notify,
  options: AutoPushOptions = {},
): Promise<boolean> {
  if (!(await savedFlag())) return false;
  await setAutoPush(true, relay, notify, options);
  return true;
}

async function apply(
  on: boolean,
  relay: RelayApi,
  notify: Notify,
  options: AutoPushOptions,
): Promise<void> {
  if (on) await arm(relay, notify, options);
  else await disarm();
  // Written only once the workbook really is watching (or really is not): a
  // failed registration must not leave a flag that re-arms on the next boot.
  await writeFlag(on);
}

async function arm(
  relay: RelayApi,
  notify: Notify,
  options: AutoPushOptions,
): Promise<void> {
  // Already watching: keep the one handler, but take this caller's relay and
  // notify - a re-installed Links tab hands over fresh ones, and pushing to the
  // client the last install created would report into a pane that is gone.
  if (handler) {
    if (session) {
      session.relay = relay;
      session.notify = notify;
    }
    return;
  }
  session = {
    relay,
    notify,
    debouncer: new Debouncer(
      options.delayMs ?? AUTOPUSH_DELAY_MS,
      flush,
      options.clock,
    ),
  };
  try {
    await Excel.run(async (context) => {
      // Committed only by the sync that registers it, like autocolor: a failed
      // sync must not leave a phantom registration behind.
      const handle = context.workbook.worksheets.onChanged.add(changed);
      await context.sync();
      handler = handle;
    });
  } catch (error) {
    session = null;
    throw error;
  }
}

async function disarm(): Promise<void> {
  const live = handler;
  if (!live) {
    forget();
    return;
  }
  // Removal runs on the context the handler was added in, and the handle is
  // given up only after it has synced, so a failed removal stays removable.
  await Excel.run(live.context, async (context) => {
    live.remove();
    await context.sync();
  });
  handler = null;
  forget();
}

function forget(): void {
  session?.debouncer.cancel();
  session = null;
  edits.clear();
}

// ---------------------------------------------------------------------------
// The edit window
// ---------------------------------------------------------------------------

// office.js hands the event a promise to wait on. The work here is deliberately
// synchronous - noting where an edit landed touches no Office.js at all - so it
// is finished before the promise this returns resolves.
function changed(event: Excel.WorksheetChangedEventArgs): Promise<void> {
  record(event);
  return Promise.resolve();
}

function record(event: Excel.WorksheetChangedEventArgs): void {
  // A removed handler still in flight stops here.
  const current = session;
  if (!current || !handler) return;
  const edit = edits.get(event.worksheetId) ?? { whole: false, areas: [] };
  if (String(event.changeType) === RANGE_EDITED) {
    // A ctrl-clicked edit carries every area it touched in one address, each
    // sheet-qualified on some hosts. The sheet is already known from the event,
    // so only the local part of each area is kept - and every area is kept, or
    // a link in the first block of a two-block paste would never be pushed.
    for (const area of splitAreas(event.address)) {
      edit.areas.push(parseAddress(area).address);
    }
  } else {
    edit.whole = true;
  }
  edits.set(event.worksheetId, edit);
  current.debouncer.touch([event.worksheetId]);
}

function flush(sheetIds: string[]): void {
  const current = session;
  if (!current) return;
  const window = new Map<string, SheetEdits>();
  for (const id of sheetIds) {
    const edit = edits.get(id);
    if (edit) {
      window.set(id, edit);
      edits.delete(id);
    }
  }
  // The registry read, the pushes and the write-back are one section, taken on
  // the queue the Links tab's own flows go through.
  void exclusive("auto-push", () => push(current, window));
}

// Nothing ever throws out of here: this runs from an event, where a rejection
// has no caller to reach.
async function push(
  current: Session,
  window: Map<string, SheetEdits>,
): Promise<void> {
  // Switched off (or re-armed) while this was queued: off means off, even for
  // a window that had already fallen due.
  if (session !== current) return;
  try {
    const ids = await affected(window);
    if (ids.length === 0) return;
    report(await pushRegistry(ids, current.relay), current.notify);
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    current.notify(`Auto-push failed: ${reason}`);
  }
}

function report(summary: PushSummary, notify: Notify): void {
  if (summary.failures.length > 0) {
    notify(`Auto-push failed: ${summary.failures.join("; ")}`);
    return;
  }
  if (summary.pushed > 0) {
    notify(`Pushed ${summary.pushed} link${summary.pushed === 1 ? "" : "s"}`);
  }
}

// Worksheet ids are what an event carries; anchors resolve to worksheet names,
// so the ids are turned into names in one batch. A sheet deleted during the
// window drops out here.
async function editedSheets(
  context: Excel.RequestContext,
  window: Map<string, SheetEdits>,
): Promise<Map<string, SheetEdits>> {
  const asked = [...window].map(([id, edit]) => ({
    sheet: context.workbook.worksheets.getItemOrNullObject(id),
    edit,
  }));
  for (const item of asked) item.sheet.load("name,isNullObject");
  await context.sync();

  const byName = new Map<string, SheetEdits>();
  for (const item of asked) {
    if (!item.sheet.isNullObject) byName.set(item.sheet.name, item.edit);
  }
  return byName;
}

function touched(entry: RegistryEntry, ref: string, edit: SheetEdits): boolean {
  if (entry.kind === "chart" || edit.whole) return true;
  return edit.areas.some((area) => intersects(ref, area));
}

async function affected(window: Map<string, SheetEdits>): Promise<string[]> {
  return Excel.run(async (context) => {
    const registry = await readRegistry(context);
    if (registry.links.length === 0) return [];
    const sheets = await editedSheets(context, window);

    // Through the anchors, never the addresses the links were created at: a
    // source may have moved since, and the edit landed where it is now. All of
    // them in one batch, because this runs every time the typing pauses.
    const resolved = await resolveSources(context, registry.links);
    const ids: string[] = [];
    registry.links.forEach((entry, index) => {
      const source = resolved[index];
      if (!source) return;
      const edit = sheets.get(source.sheet);
      if (edit && touched(entry, source.ref, edit)) ids.push(entry.id);
    });
    return ids;
  });
}

// ---------------------------------------------------------------------------
// The workbook flag
// ---------------------------------------------------------------------------

async function writeFlag(on: boolean): Promise<void> {
  await Excel.run(async (context) => {
    context.workbook.settings.add(AUTOPUSH_SETTING, on ? ENABLED : DISABLED);
    await context.sync();
  });
}

async function savedFlag(): Promise<boolean> {
  return Excel.run(async (context) => {
    const setting =
      context.workbook.settings.getItemOrNullObject(AUTOPUSH_SETTING);
    setting.load("isNullObject,value");
    await context.sync();
    return !setting.isNullObject && String(setting.value) === ENABLED;
  });
}
