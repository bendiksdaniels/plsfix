// A hand-built PowerPoint host: an in-memory deck answering exactly the
// office.js surface src/ppt/host.ts touches, so link insert, refresh and break
// can be driven without a sideload. Strict load semantics are opt-in and
// global, the way the Excel fake does it; the two fakes share no code.

import type { FakePresentation } from "./model";
import { installGlobals, removeGlobals } from "./office";

export type {
  FakePptShape,
  FakeShapeGroup,
  FakeShapeInit,
  FakeSlide,
  ShapeSite,
} from "./model";
export { FakePresentation } from "./model";

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
  setSupported(check: (set: string, version: string) => boolean): void;
  storage(): Map<string, string>;
  insertedViaSelection(): SelectionInsert[];
  // Makes the next setSelectedDataAsync report failure, the way a host that is
  // out of memory or has the slide locked does. Consumed by that one call.
  failNextSelectionInsert(message?: string): void;
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
