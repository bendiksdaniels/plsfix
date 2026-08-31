// The host itself: the runtime a test installs, the request context PowerPoint
// hands a batch, and the PowerPoint, Office and OfficeRuntime globals the pane
// talks to. The slide, shape and tag objects the context leads to live in
// objects.ts, over the deck in model.ts.

import type { Box, FakePptHelpers, FakePptOptions, SelectionInsert } from ".";
import { FakePresentation } from "./model";
import { SlideCollectionProxy } from "./objects";
import { Loadable, StrictLoads } from "./strict";

interface SelectionOptions {
  coercionType?: string;
  imageLeft?: number;
  imageTop?: number;
  imageWidth?: number;
  imageHeight?: number;
}

class FakeRuntime {
  strict: StrictLoads | null;
  supported: (set: string, version: string) => boolean;
  storage: Map<string, string>;
  insertions: SelectionInsert[] = [];
  nextInsertFailure: string | null = null;
  // Office.context.platform: the desktop until a test says otherwise, because
  // the web is the host with the tighter shape budget.
  platform = "Mac";
  // Every context.sync() this host served: one round trip to PowerPoint, and
  // the only cost a deck's size is allowed to multiply.
  syncs = 0;

  constructor(
    public presentation: FakePresentation,
    options: FakePptOptions,
    strictDefault: boolean,
  ) {
    this.supported = options.isSetSupported ?? (() => true);
    this.storage = options.storage ?? new Map();
    this.strict =
      (options.strictLoad ?? strictDefault) ? new StrictLoads() : null;
  }
}

class FakeContext extends Loadable {
  presentation: PresentationProxy;

  constructor(private runtime: FakeRuntime) {
    super();
    this.presentation = new PresentationProxy(runtime.presentation);
  }

  // A no-op flush: reads come from the deck, so only load state moves here.
  // Counted, because a batch that syncs per shape is the performance bug.
  // Every shape added in the batch is now one the host has heard of.
  sync(): Promise<void> {
    this.runtime.syncs += 1;
    this.runtime.presentation.markSynced();
    this.runtime.strict?.commit();
    return Promise.resolve();
  }
}

class PresentationProxy extends Loadable {
  constructor(private deck: FakePresentation) {
    super();
  }

  get slides(): SlideCollectionProxy {
    return new SlideCollectionProxy(this.deck, () => this.deck.slides);
  }
  getSelectedSlides(): SlideCollectionProxy {
    return new SlideCollectionProxy(this.deck, () =>
      this.deck.selectedSlides(),
    );
  }
  setSelectedSlides(slideIds: string[]): void {
    this.deck.selectedSlideIds = [...slideIds];
  }
}

// ---------------------------------------------------------------------------
// Globals
// ---------------------------------------------------------------------------

function pickCallback(
  ...candidates: unknown[]
): ((result: unknown) => void) | undefined {
  for (const candidate of candidates) {
    if (typeof candidate === "function") {
      return candidate as (result: unknown) => void;
    }
  }
  return undefined;
}

// Inserting a picture through the selection API, the only route on hosts
// without the shape APIs: it lands on the selected slide, at the given box.
function insertViaSelection(
  runtime: FakeRuntime,
  png: string,
  options: SelectionOptions,
): void {
  const failure = runtime.nextInsertFailure;
  if (failure !== null) {
    runtime.nextInsertFailure = null;
    throw new Error(failure);
  }
  const deck = runtime.presentation;
  const slideId = deck.selectedSlideIds[0] ?? deck.slides[0]?.id ?? "";
  const slide = deck.findSlideOrThrow(slideId);
  const box: Box = {
    left: options.imageLeft ?? 0,
    top: options.imageTop ?? 0,
    width: options.imageWidth ?? 0,
    height: options.imageHeight ?? 0,
  };
  // The picture is the shape's own image; nothing called fill.setImage.
  deck.addShape(slide, { type: "Image", fillImage: png, ...box });
  runtime.insertions.push({ slideId: slide.id, png, box });
}

function documentApi(runtime: FakeRuntime): Record<string, unknown> {
  return {
    setSelectedDataAsync(data: unknown, first?: unknown, second?: unknown) {
      const done = pickCallback(first, second);
      const options = (
        typeof first === "object" && first !== null ? first : {}
      ) as SelectionOptions;
      queueMicrotask(() => {
        try {
          insertViaSelection(runtime, String(data), options);
          done?.({ status: "succeeded", value: undefined });
        } catch (error) {
          done?.({ status: "failed", error });
        }
      });
    },
    getFilePropertiesAsync(first?: unknown, second?: unknown) {
      const done = pickCallback(first, second);
      const url = runtime.presentation.fileUrl;
      queueMicrotask(() => {
        done?.({ status: "succeeded", value: { url } });
      });
    },
  };
}

function officeGlobal(runtime: FakeRuntime): Record<string, unknown> {
  return {
    context: {
      get platform(): string {
        return runtime.platform;
      },
      requirements: {
        isSetSupported: (set: string, version: string) =>
          runtime.supported(set, version),
      },
      document: documentApi(runtime),
    },
    HostType: { Excel: "Excel", Word: "Word", PowerPoint: "PowerPoint" },
    // office.js gives these string values, whatever the .d.ts enum looks like.
    PlatformType: {
      PC: "PC",
      OfficeOnline: "OfficeOnline",
      Mac: "Mac",
      iOS: "iOS",
      Android: "Android",
      Universal: "Universal",
    },
    CoercionType: { Text: "text", Image: "image", SlideRange: "slideRange" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    // Office.onReady both calls back and resolves with the host it found.
    onReady: (callback?: (info: { host: string }) => unknown) => {
      const info = { host: "PowerPoint" };
      callback?.(info);
      return Promise.resolve(info);
    },
  };
}

function powerPointGlobal(runtime: FakeRuntime): Record<string, unknown> {
  return {
    run(first: unknown, second?: unknown): Promise<unknown> {
      const callback = (typeof first === "function" ? first : second) as (
        context: unknown,
      ) => unknown;
      const context = new FakeContext(runtime);
      const handed = runtime.strict?.root(context, "context") ?? context;
      return Promise.resolve().then(() => callback(handed));
    },
    GeometricShapeType: {
      rectangle: "Rectangle",
      ellipse: "Ellipse",
      pie: "Pie",
    },
    ConnectorType: { straight: "Straight", elbow: "Elbow", curve: "Curve" },
    ShapeAutoSize: {
      autoSizeNone: "AutoSizeNone",
      autoSizeTextToFitShape: "AutoSizeTextToFitShape",
      autoSizeShapeToFitText: "AutoSizeShapeToFitText",
      autoSizeMixed: "AutoSizeMixed",
    },
    ShapeType: {
      unsupported: "Unsupported",
      image: "Image",
      geometricShape: "GeometricShape",
      group: "Group",
      line: "Line",
      table: "Table",
    },
  };
}

function officeRuntimeGlobal(runtime: FakeRuntime): Record<string, unknown> {
  return {
    storage: {
      getItem: (key: string) =>
        Promise.resolve(runtime.storage.get(key) ?? null),
      setItem: (key: string, value: string) => {
        runtime.storage.set(key, value);
        return Promise.resolve();
      },
      removeItem: (key: string) => {
        runtime.storage.delete(key);
        return Promise.resolve();
      },
    },
  };
}

function makeHelpers(runtime: FakeRuntime): FakePptHelpers {
  return {
    selectSlide(id) {
      runtime.presentation.findSlideOrThrow(id);
      runtime.presentation.selectedSlideIds = [id];
    },
    clearSelection() {
      runtime.presentation.selectedSlideIds = [];
    },
    setSupported(check) {
      runtime.supported = check;
    },
    setPlatform(platform) {
      runtime.platform = platform;
    },
    storage: () => runtime.storage,
    syncCount: () => runtime.syncs,
    failNextSelectionInsert(
      message = "PowerPoint could not insert the image.",
    ) {
      runtime.nextInsertFailure = message;
    },
    insertedViaSelection: () =>
      runtime.insertions.map((insert) => ({
        ...insert,
        box: { ...insert.box },
      })),
  };
}

export function installGlobals(
  options: FakePptOptions,
  strictDefault: boolean,
): { presentation: FakePresentation; helpers: FakePptHelpers } {
  const presentation =
    options.presentation ?? new FakePresentation(options.slides);
  const runtime = new FakeRuntime(presentation, options, strictDefault);
  const scope = globalThis as unknown as Record<string, unknown>;
  scope.PowerPoint = powerPointGlobal(runtime);
  scope.Office = officeGlobal(runtime);
  scope.OfficeRuntime = officeRuntimeGlobal(runtime);
  return { presentation, helpers: makeHelpers(runtime) };
}

export function removeGlobals(): void {
  const scope = globalThis as unknown as Record<string, unknown>;
  delete scope.PowerPoint;
  delete scope.Office;
  delete scope.OfficeRuntime;
}
