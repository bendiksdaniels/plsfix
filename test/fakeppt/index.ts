// A hand-built PowerPoint host: an in-memory deck answering exactly the
// office.js surface src/ppt/host.ts touches, so link insert, refresh and break
// can be driven without a sideload. Strict load semantics are opt-in and
// global, the way the Excel fake does it; the two fakes share no code.

import type { FakePresentation } from "./model";
import { installGlobals, removeGlobals } from "./office";

export type {
  FakePptShape,
  FakeShapeFont,
  FakeShapeGroup,
  FakeShapeInit,
  FakeShapeMargins,
  FakeSlide,
  ShapeSite,
} from "./model";
export type {
  FakeTable,
  FakeTableCell,
  FakeTableCellFill,
  FakeTableStyleSettings,
} from "./tables";
export { FakePresentation } from "./model";
// The current selection, in order: what getSelectedShapes() hands back.
export { selectedShapeIds } from "./objects";

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

// One Office.context.document.setSelectedDataAsync call, as the fake saw it.
export interface SelectionInsert {
  slideId: string;
  png: string;
  box: Box;
}

export interface FakePptOptions {
  presentation?: FakePresentation;
  slides?: number;
  isSetSupported?: (set: string, version: string) => boolean;
  strictLoad?: boolean;
  storage?: Map<string, string>;
}

export interface FakePptHelpers {
  selectSlide(id: string): void;
  clearSelection(): void;
  // Sets the shape selection directly, order preserved, skipping the
  // same-slide rule Slide.setSelectedShapes enforces: a fixture shortcut
  // mirroring selectSlide.
  selectShapes(shapeIds: string[]): void;
  setSupported(check: (set: string, version: string) => boolean): void;
  // Registers a style the NEXT shapes.addTable on this shape id starts with,
  // in place of PowerPoint's own default (Medium Style 2 Accent 1) - the only
  // way to make a just-inserted or just-rebuilt table arrive with no style,
  // or with one already chosen. Ids are sequential per deck (`shape-1`,
  // `shape-2`, ...), so a fresh presentation's first table is always
  // "shape-1".
  setTableStyle(shapeId: string, style: string): void;
  // Office.context.platform, "Mac" until a test asks for another host: the web
  // is the one that draws fewer shapes before it gives up.
  setPlatform(platform: string): void;
  storage(): Map<string, string>;
  // Round trips to the host since this fake was installed: what a batching
  // change is measured in, counted from zero per installFakePpt.
  syncCount(): number;
  insertedViaSelection(): SelectionInsert[];
  // Makes the next setSelectedDataAsync report failure, the way a host that is
  // out of memory or has the slide locked does. Consumed by that one call.
  failNextSelectionInsert(message?: string): void;
  // Makes a future context.sync() reject, the way a hung or refused round
  // trip does: the very next one by default, or the one after skipping
  // `afterSyncs` more that still succeed. Calls stack, so more than one
  // future sync can be armed to fail. Whichever sync it lands on takes back
  // off the deck everything added since the sync before it.
  failNextSync(error?: Error, afterSyncs?: number): void;
  // Makes a future context.sync() reject the way a real refusal does: the
  // host applied the batch's adds and then refused, so they stay on the deck
  // with ids no load answered (Mac 16.107, 14.09: four sub-groups landed and
  // the sync rejected on the fifth addGroup). Same ordinal rule as
  // failNextSync; every add of that batch is kept, the loads are lost.
  refuseNextSync(error?: Error, afterSyncs?: number): void;
  // G audit: makes a future context.sync() never settle - neither applying
  // its batch nor rejecting it - the way PowerPoint for the web swallowed a
  // chart draw's second chunk (tasks/lessons.md, 2026-09-08). Same ordinal
  // rule as failNextSync; the shapes of that batch never reach the deck.
  hangNextSync(afterSyncs?: number): void;
  // The ribbon FunctionNames registerCommands associated with this host.
  commandIds(): string[];
  // Presses one of them the way the ribbon does, resolving when the command
  // calls event.completed().
  runCommand(id: string): Promise<void>;
}

let strictByDefault = false;

// Switches strict load semantics on for every host installed afterwards, which
// is how a suite runs the add-in against real office.js load ordering.
// installFakePpt({ strictLoad }) overrides it per host.
export function enableStrictLoadSemantics(on = true): void {
  strictByDefault = on;
}

export function installFakePpt(options: FakePptOptions = {}): {
  presentation: FakePresentation;
  helpers: FakePptHelpers;
} {
  return installGlobals(options, strictByDefault);
}

export function uninstallFakePpt(): void {
  removeGlobals();
}
