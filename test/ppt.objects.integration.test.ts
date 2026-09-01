// PowerPoint object tools (align, distribute, match size, swap, select
// similar, Smart Painter) against the strict fake host: every geometry mode,
// every count refusal with its exact message, the PowerPointApi 1.5 guard,
// and the sync budget - two round trips per geometry transform, whatever the
// selection size. Strict load semantics are on, so a missing load() fails
// here first.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  selectedShapeIds,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
  type FakeShapeInit,
  type FakeSlide,
} from "./fakeppt";
import type { AlignMode } from "../src/ppt/object-math";
import type * as ObjectToolsModule from "../src/ppt/object-tools";

enableStrictLoadSemantics();

let tools: typeof ObjectToolsModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 2 });
  presentation = host.presentation;
  helpers = host.helpers;
  tools = await import("../src/ppt/object-tools");
});
afterEach(() => {
  uninstallFakePpt();
});

function box(
  left: number,
  top: number,
  width: number,
  height: number,
): FakeShapeInit {
  return { left, top, width, height };
}

// The same three boxes object-math.test.ts checks the pure maths against, so
// the expected numbers here are the ones already proven there - this suite
// only has to prove the office.js round trip reaches the same result.
function seedThree(
  slide: FakeSlide = presentation.slides[0]!,
): [string, string, string] {
  const a = presentation.addShape(slide, box(10, 20, 30, 20));
  const b = presentation.addShape(slide, box(60, 50, 20, 30));
  const c = presentation.addShape(slide, box(110, 100, 40, 10));
  return [a.id, b.id, c.id];
}

function positions(
  ids: string[],
): { left: number; top: number; width: number; height: number }[] {
  return ids.map((id) => {
    const { shape } = presentation.findShape(id);
    return {
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
    };
  });
}

describe("align", () => {
  const cases: [AlignMode, "left" | "top", number[]][] = [
    ["left", "left", [10, 10, 10]],
    ["center", "left", [65, 70, 60]],
    ["right", "left", [120, 130, 110]],
    ["top", "top", [20, 20, 20]],
    ["middle", "top", [55, 50, 60]],
    ["bottom", "top", [90, 80, 100]],
  ];

  it.each(cases)(
    "aligns %s against the selection bounds",
    async (mode, axis, expected) => {
      const ids = seedThree();
      helpers.selectShapes(ids);
      const message = await tools.alignSelected(mode);
      expect(message).toBe(`Aligned 3 objects ${mode}.`);
      expect(positions(ids).map((p) => p[axis])).toEqual(expected);
    },
  );

  it("refuses fewer than two objects", async () => {
    const [aId] = seedThree();
    helpers.selectShapes([aId]);
    await expect(tools.alignSelected("left")).rejects.toThrow(
      "Select at least two objects to align.",
    );
  });
});

describe("distribute", () => {
  it("distributes across, keeping the outer edges fixed", async () => {
    const ids = seedThree();
    helpers.selectShapes(ids);
    const message = await tools.distributeSelected("horizontal");
    expect(message).toBe("Distributed 3 objects across.");
    expect(positions(ids).map((p) => p.left)).toEqual([10, 65, 110]);
  });

  it("distributes down, keeping the outer edges fixed", async () => {
    const ids = seedThree();
    helpers.selectShapes(ids);
    const message = await tools.distributeSelected("vertical");
    expect(message).toBe("Distributed 3 objects down.");
    expect(positions(ids).map((p) => p.top)).toEqual([20, 55, 100]);
  });

  it("refuses fewer than three objects", async () => {
    const [aId, bId] = seedThree();
    helpers.selectShapes([aId, bId]);
    await expect(tools.distributeSelected("horizontal")).rejects.toThrow(
      "Select at least three objects to distribute.",
    );
  });
});

describe("match size", () => {
  it("matches targets to the first shape's size, leaving it unchanged", async () => {
    const ids = seedThree();
    helpers.selectShapes(ids);
    const message = await tools.matchSelectedSize();
    expect(message).toBe("Matched 2 objects to the reference size.");
    const [a, b, c] = positions(ids);
    expect(a).toEqual({ left: 10, top: 20, width: 30, height: 20 });
    expect(b).toMatchObject({ width: 30, height: 20 });
    expect(c).toMatchObject({ width: 30, height: 20 });
  });

  it("refuses fewer than two objects", async () => {
    const [aId] = seedThree();
    helpers.selectShapes([aId]);
    await expect(tools.matchSelectedSize()).rejects.toThrow(
      "Select the reference object first, then at least one target.",
    );
  });
});

describe("swap", () => {
  it("swaps the positions of exactly two objects", async () => {
    const [aId, bId] = seedThree();
    helpers.selectShapes([aId, bId]);
    const message = await tools.swapSelected();
    expect(message).toBe("Swapped two objects.");
    const [a, b] = positions([aId, bId]);
    expect(a).toMatchObject({ left: 60, top: 50 });
    expect(b).toMatchObject({ left: 10, top: 20 });
  });

  it("refuses one object", async () => {
    const [aId] = seedThree();
    helpers.selectShapes([aId]);
    await expect(tools.swapSelected()).rejects.toThrow(
      "Select exactly two objects to swap.",
    );
  });

  it("refuses three objects", async () => {
    const ids = seedThree();
    helpers.selectShapes(ids);
    await expect(tools.swapSelected()).rejects.toThrow(
      "Select exactly two objects to swap.",
    );
  });
});

describe("select similar", () => {
  it("matches same type and size within 1pt, same slide only, keeping the reference selected", async () => {
    const [slide0, slide1] = presentation.slides as [FakeSlide, FakeSlide];
    const reference = presentation.addShape(slide0, box(0, 0, 100, 50));
    const closeMatch = presentation.addShape(slide0, box(200, 0, 100.5, 50));
    const wrongType = presentation.addShape(slide0, {
      ...box(300, 0, 100, 50),
      type: "TextBox",
    });
    const tooDifferent = presentation.addShape(slide0, box(400, 0, 106, 50));
    presentation.addShape(slide1, box(0, 0, 100, 50));

    helpers.selectShapes([reference.id]);
    const message = await tools.selectSimilar();
    expect(message).toBe("Selected 2 similar objects.");
    expect(selectedShapeIds(presentation)).toEqual([
      reference.id,
      closeMatch.id,
    ]);
    expect(selectedShapeIds(presentation)).not.toContain(wrongType.id);
    expect(selectedShapeIds(presentation)).not.toContain(tooDifferent.id);
  });

  it("refuses more than one reference object", async () => {
    const ids = seedThree();
    helpers.selectShapes(ids);
    await expect(tools.selectSimilar()).rejects.toThrow(
      "Select one reference object.",
    );
  });
});

describe("Smart Painter", () => {
  it("captures and applies a solid fill and line format", async () => {
    const slide = presentation.slides[0]!;
    const source = presentation.addShape(slide, box(0, 0, 40, 40));
    source.fillColor = "#B27E54";
    source.lineColor = "#282623";
    source.lineWeight = 1.5;
    const target = presentation.addShape(slide, box(100, 0, 20, 20));

    helpers.selectShapes([source.id]);
    expect(await tools.captureObjectStyle()).toBe("Object style captured.");
    helpers.selectShapes([target.id]);
    expect(await tools.applyObjectStyle()).toBe("Painted 1 object.");

    const { shape } = presentation.findShape(target.id);
    expect(shape.fillColor).toBe("#B27E54");
    expect(shape.fillCleared).toBe(false);
    expect(shape.lineColor).toBe("#282623");
    expect(shape.lineWeight).toBe(1.5);
  });

  it("captures and applies a no-fill style", async () => {
    const slide = presentation.slides[0]!;
    const source = presentation.addShape(slide, box(0, 0, 40, 40));
    const target = presentation.addShape(slide, box(100, 0, 20, 20));
    target.fillColor = "#FFFFFF";

    helpers.selectShapes([source.id]);
    await tools.captureObjectStyle();
    helpers.selectShapes([target.id]);
    await tools.applyObjectStyle();

    const { shape } = presentation.findShape(target.id);
    expect(shape.fillColor).toBeNull();
    expect(shape.fillCleared).toBe(true);
  });

  it("refuses to capture a picture fill", async () => {
    const slide = presentation.slides[0]!;
    const picture = presentation.addShape(slide, {
      ...box(0, 0, 40, 40),
      fillImage: "<png>",
    });
    helpers.selectShapes([picture.id]);
    await expect(tools.captureObjectStyle()).rejects.toThrow(
      "Smart Painter currently supports solid or no-fill objects.",
    );
  });

  it("refuses to apply before anything was captured", async () => {
    const [aId] = seedThree();
    helpers.selectShapes([aId]);
    await expect(tools.applyObjectStyle()).rejects.toThrow(
      "Capture an object style first.",
    );
  });

  it("refuses an empty selection once a style is captured", async () => {
    const [aId] = seedThree();
    helpers.selectShapes([aId]);
    await tools.captureObjectStyle();
    helpers.selectShapes([]);
    await expect(tools.applyObjectStyle()).rejects.toThrow(
      "Select at least one target object.",
    );
  });

  it("refuses to capture more than one object", async () => {
    const ids = seedThree();
    helpers.selectShapes(ids);
    await expect(tools.captureObjectStyle()).rejects.toThrow(
      "Select one object to capture.",
    );
  });
});

describe("sync budget", () => {
  it("spends exactly two syncs per geometry transform", async () => {
    const ids = seedThree();

    helpers.selectShapes(ids);
    let from = helpers.syncCount();
    await tools.alignSelected("left");
    expect(helpers.syncCount() - from).toBe(2);

    helpers.selectShapes(ids);
    from = helpers.syncCount();
    await tools.distributeSelected("horizontal");
    expect(helpers.syncCount() - from).toBe(2);

    helpers.selectShapes(ids);
    from = helpers.syncCount();
    await tools.matchSelectedSize();
    expect(helpers.syncCount() - from).toBe(2);

    helpers.selectShapes(ids.slice(0, 2));
    from = helpers.syncCount();
    await tools.swapSelected();
    expect(helpers.syncCount() - from).toBe(2);
  });

  it("spends three syncs to select similar, two to capture, two to apply", async () => {
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, box(0, 0, 20, 20));
    const b = presentation.addShape(slide, box(50, 0, 20, 20));

    helpers.selectShapes([a.id]);
    let from = helpers.syncCount();
    await tools.selectSimilar();
    expect(helpers.syncCount() - from).toBe(3);

    helpers.selectShapes([a.id]);
    from = helpers.syncCount();
    await tools.captureObjectStyle();
    expect(helpers.syncCount() - from).toBe(2);

    helpers.selectShapes([b.id]);
    from = helpers.syncCount();
    await tools.applyObjectStyle();
    expect(helpers.syncCount() - from).toBe(2);
  });
});

describe("PowerPointApi 1.5 guard", () => {
  it("refuses every object tool on an older host, before any sync", async () => {
    helpers.setSupported((_set, version) => version !== "1.5");
    const ids = seedThree();
    helpers.selectShapes(ids);
    const from = helpers.syncCount();
    const message = "Object tools need PowerPoint 2021 or Microsoft 365.";
    await expect(tools.alignSelected("left")).rejects.toThrow(message);
    await expect(tools.selectSimilar()).rejects.toThrow(message);
    await expect(tools.captureObjectStyle()).rejects.toThrow(message);
    expect(helpers.syncCount()).toBe(from);
  });
});

describe("strict load semantics", () => {
  it("catches a read of an unloaded fill scalar", async () => {
    const slide = presentation.slides[0]!;
    const shape = presentation.addShape(slide, box(0, 0, 10, 10));
    await expect(
      PowerPoint.run((c) => {
        const target = c.presentation.slides
          .getItemAt(0)
          .shapes.getItem(shape.id);
        return Promise.resolve(target.fill.type);
      }),
    ).rejects.toThrow(/PropertyNotLoaded/);
  });
});
