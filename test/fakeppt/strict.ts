// Real office.js hands out proxies, not values: reading a scalar nobody asked
// for with load() and committed with sync() throws. The fake serves reads from
// the deck, so this layer is what keeps load ordering honest. Load state hangs
// off the object load() was called on (the ROOT), keyed by property path.

const RAW = Symbol("fakeppt.raw");

interface Descriptor {
  // Property reads that need a load plus a sync first.
  scalars?: readonly string[];
  // Navigation property -> kind of the child it hands back; no load of its own.
  children?: Readonly<Record<string, string>>;
  // Method -> kind of the fresh, nothing-loaded root it hands back.
  returns?: Readonly<Record<string, string>>;
  // Kind of the elements behind a collection's items array.
  items?: string;
  // ClientResult: .value arrives with the next sync, no load needed.
  result?: boolean;
}

// Every kind the add-in can reach. A property named nowhere here goes
// unpoliced, so the scalar lists carry the office.js surface in full.
const KINDS: Record<string, Descriptor> = {
  context: { children: { presentation: "presentation" } },
  presentation: {
    children: { slides: "slides" },
    returns: { getSelectedSlides: "slides" },
  },
  slides: {
    items: "slide",
    returns: {
      getItem: "slide",
      getItemAt: "slide",
      getItemOrNullObject: "slide",
      getCount: "clientResult",
    },
  },
  slide: {
    scalars: ["id", "index", "isNullObject"],
    children: { shapes: "shapes" },
  },
  shapes: {
    items: "shape",
    returns: {
      getItem: "shape",
      getItemOrNullObject: "shape",
      addGeometricShape: "shape",
      addTextBox: "shape",
      addLine: "shape",
      addGroup: "shape",
      addTable: "shape",
      getCount: "clientResult",
    },
  },
  shape: {
    scalars: [
      "id",
      "name",
      "type",
      "left",
      "top",
      "width",
      "height",
      "zOrderPosition",
      "isNullObject",
    ],
    children: {
      fill: "fill",
      lineFormat: "lineFormat",
      tags: "tags",
      group: "group",
      textFrame: "textFrame",
      adjustments: "adjustments",
    },
    returns: { getParentSlideOrNullObject: "slide", getTable: "table" },
  },
  // A shape of type Group and the shapes inside it: one more level of the same
  // items/<property> load paths, which is how office.js reads a group too.
  group: { scalars: ["id"], children: { shapes: "groupShapes" } },
  groupShapes: {
    items: "shape",
    returns: {
      getItem: "shape",
      getItemOrNullObject: "shape",
      getCount: "clientResult",
    },
  },
  fill: {},
  lineFormat: { scalars: ["visible", "color", "weight"] },
  // The geometry's own handles: a pie's start and end angle, read back one at
  // a time through a ClientResult.
  adjustments: { scalars: ["count"], returns: { get: "clientResult" } },
  textFrame: {
    scalars: [
      "hasText",
      "autoSizeSetting",
      "wordWrap",
      "verticalAlignment",
      "leftMargin",
      "rightMargin",
      "topMargin",
      "bottomMargin",
    ],
    children: { textRange: "textRange" },
  },
  textRange: {
    scalars: ["text"],
    children: { font: "textFont", paragraphFormat: "paragraphFormat" },
  },
  textFont: { scalars: ["name", "size", "color", "bold"] },
  paragraphFormat: { scalars: ["horizontalAlignment"] },
  // A native table and one of its cells: the counts are read to decide whether
  // a repaint fits, everything else is written.
  table: {
    scalars: ["rowCount", "columnCount"],
    returns: { getCellOrNullObject: "tableCell" },
  },
  tableCell: {
    scalars: ["text", "horizontalAlignment"],
    children: { font: "tableCellFont", fill: "tableCellFill" },
  },
  tableCellFont: { scalars: ["bold", "italic", "color", "size"] },
  tableCellFill: {},
  tags: { items: "tag", returns: { getItemOrNullObject: "tag" } },
  tag: { scalars: ["key", "value", "isNullObject"] },
  clientResult: { result: true },
};

export function notLoaded(property: string): Error {
  const error = new Error(`PropertyNotLoaded: ${property}`) as Error & {
    code: string;
  };
  error.code = "PropertyNotLoaded";
  return error;
}

// What getCount() hands back: a value that only exists after the next sync.
export class FakeClientResult<T> {
  constructor(readonly value: T) {}
}

// Every office.js object answers load(); the strict layer intercepts the call
// below, so this body only runs when strict semantics are off.
export abstract class Loadable {
  load(): this {
    return this;
  }
}

interface LoadState {
  loaded: Set<string>;
  pending: Set<string>;
  // Prefixes a bare load() covered: every direct scalar below them.
  all: Set<string>;
  pendingAll: Set<string>;
  result: boolean;
  pendingResult: boolean;
}

function loadPaths(argument: unknown): string[] | null {
  const split = (text: string): string[] =>
    text
      .split(",")
      .map((part) => part.trim())
      .filter(Boolean);

  if (typeof argument === "string") return split(argument);
  if (Array.isArray(argument)) {
    return (argument as unknown[]).flatMap((entry) => split(String(entry)));
  }
  if (argument !== null && typeof argument === "object") {
    return loadPaths((argument as { select?: unknown }).select);
  }
  // load() with no argument: every scalar of that object.
  return null;
}

function rawOf<T>(value: T): T {
  if (value !== null && typeof value === "object") {
    const raw = (value as Record<symbol, unknown>)[RAW];
    if (raw !== undefined) return raw as T;
  }
  return value;
}

export class StrictLoads {
  private states = new WeakMap<object, LoadState>();
  private queued = new Set<LoadState>();

  // A sync is what makes a requested property readable.
  commit(): void {
    for (const state of this.queued) {
      for (const path of state.pending) state.loaded.add(path);
      for (const prefix of state.pendingAll) state.all.add(prefix);
      if (state.pendingResult) state.result = true;
      state.pending.clear();
      state.pendingAll.clear();
      state.pendingResult = false;
    }
    this.queued.clear();
  }

  // Wraps an object owning its own load state: the context, and everything a
  // method hands back.
  root<T>(value: T, kind: string): T {
    if (value === null || typeof value !== "object") return value;
    if (KINDS[kind]?.result) {
      const state = this.state(value);
      state.pendingResult = true;
      this.queued.add(state);
    }
    return this.wrap(value, kind, value, "") as T;
  }

  private state(root: object): LoadState {
    let found = this.states.get(root);
    if (!found) {
      found = {
        loaded: new Set(),
        pending: new Set(),
        all: new Set(),
        pendingAll: new Set(),
        result: false,
        pendingResult: false,
      };
      this.states.set(root, found);
    }
    return found;
  }

  private record(root: object, prefix: string, argument: unknown): void {
    const state = this.state(root);
    const paths = loadPaths(argument);
    if (paths === null) state.pendingAll.add(prefix);
    else for (const path of paths) state.pending.add(prefix + path);
    this.queued.add(state);
  }

  private require(
    root: object,
    path: string,
    property: string,
    collection: boolean,
  ): void {
    const state = this.states.get(root);
    if (state) {
      if (state.loaded.has(path)) return;
      const cut = path.lastIndexOf("/");
      if (state.all.has(cut < 0 ? "" : path.slice(0, cut + 1))) return;
      // "items/id" makes the items array readable, as it does in PowerPoint.
      if (collection) {
        for (const loaded of state.loaded) {
          if (loaded.startsWith(`${path}/`)) return;
        }
      }
    }
    throw notLoaded(property);
  }

  private wrap(
    value: unknown,
    kind: string,
    root: object,
    prefix: string,
  ): unknown {
    if (value === null || typeof value !== "object") return value;
    const descriptor = KINDS[kind];
    // A kind nobody described would police nothing, which is the one failure
    // this layer must never have: say so instead of passing everything through.
    if (!descriptor) throw new Error(`fake ppt has no strict kind "${kind}"`);

    return new Proxy(value as Record<string, unknown>, {
      get: (target, property, receiver) => {
        if (property === RAW) return target;
        if (typeof property === "symbol") return Reflect.get(target, property);
        return this.read(target, property, descriptor, root, prefix, receiver);
      },
      set(target, property, next) {
        // Writes never need a load, and the setter runs against the deck.
        return Reflect.set(target, property, next);
      },
    });
  }

  private read(
    target: Record<string, unknown>,
    property: string,
    descriptor: Descriptor,
    root: object,
    prefix: string,
    receiver: unknown,
  ): unknown {
    const path = prefix + property;
    const childKind = descriptor.children?.[property];
    if (childKind !== undefined) {
      const child: unknown = Reflect.get(target, property);
      return this.wrap(child, childKind, root, `${path}/`);
    }
    const itemKind = descriptor.items;
    if (property === "items" && itemKind !== undefined) {
      this.require(root, path, property, true);
      const items = Reflect.get(target, property) as unknown[];
      return items.map((item) => this.wrap(item, itemKind, root, `${path}/`));
    }
    if (descriptor.scalars?.includes(property)) {
      this.require(root, path, property, false);
      return Reflect.get(target, property);
    }
    if (property === "value" && descriptor.result) {
      if (!this.state(root).result) throw notLoaded(property);
      return Reflect.get(target, property);
    }
    const raw: unknown = Reflect.get(target, property);
    if (typeof raw === "function") {
      return this.method(target, property, descriptor, root, prefix, receiver);
    }
    // The fake's own fields, and the "then" the runtime probes for whenever a
    // proxy comes back from an async function.
    return raw;
  }

  private method(
    target: Record<string, unknown>,
    property: string,
    descriptor: Descriptor,
    root: object,
    prefix: string,
    self: unknown,
  ): (...args: unknown[]) => unknown {
    return (...args: unknown[]): unknown => {
      // Host-side calls take objects, not property reads: hand the fake its own
      // unwrapped proxies so its internals are not policed as add-in reads.
      const plain = args.map((argument) => rawOf(argument));
      if (property === "load") {
        this.record(root, prefix, plain[0]);
        return self;
      }
      const call = target[property] as (...rest: unknown[]) => unknown;
      const result: unknown = call.apply(target, plain);
      const kind = descriptor.returns?.[property];
      return kind === undefined ? result : this.root(result, kind);
    };
  }
}
