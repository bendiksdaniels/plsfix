// Stress pass on the PowerPoint Tools tab: every object tool driven with the
// wrong selection (nothing, a group's child, a placeholder), against a host
// that refuses or swallows the write, and Smart Painter over the shapes a deck
// actually holds. Invariant: every abuse ends in one sentence, and no tool
// reports a success it did not perform.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { settleHungSync } from "./hung-sync";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  selectedShapeIds,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
  type FakeShapeInit,
} from "./fakeppt";
import type * as ToolsModule from "../src/ppt/object-tools";

enableStrictLoadSemantics();

let tools: typeof ToolsModule;
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
  vi.useRealTimers();
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

function add(init: FakeShapeInit, slide = 0): FakePptShape {
  return presentation.addShape(presentation.slides[slide]!, init);
}

function two(): [FakePptShape, FakePptShape] {
  return [add(box(10, 20, 30, 20)), add(box(60, 50, 20, 30))];
}

// The four numbers every geometry tool writes, and the only state a refusal
// is allowed to leave behind unchanged.
function boxesOf(shapes: FakePptShape[]): FakeShapeInit[] {
  return shapes.map(({ left, top, width, height }) => ({
    left,
    top,
    width,
    height,
  }));
}

// Every tool the Tools tab and the ribbon share, by the name its button wears.
function everyTool(): [string, () => Promise<string>][] {
  return [
    ["Align", () => tools.alignSelected("left")],
    ["Distribute", () => tools.distributeSelected("horizontal")],
    ["Match size", () => tools.matchSelectedSize()],
    ["Swap", () => tools.swapSelected()],
    ["Select similar", () => tools.selectSimilar()],
    ["Capture style", () => tools.captureObjectStyle()],
    ["Apply style", () => tools.applyObjectStyle()],
  ];
}

async function refusal(run: () => Promise<string>): Promise<string> {
  try {
    const message = await run();
    throw new Error(`expected a refusal, got "${message}"`);
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
}

describe("a tool pressed with nothing selected", () => {
  it.each(everyTool())(
    "%s asks for a selection in one sentence, and spends no write",
    async (_name, run) => {
      const shapes = two();
      helpers.selectShapes([]);
      const before = boxesOf(shapes);

      const message = await refusal(run);

      // A sentence: it ends in a full stop and names no office.js code.
      expect(message).toMatch(/\.$/);
      expect(message).not.toMatch(/Exception|NotFound|InvalidArgument|null/);
      expect(boxesOf(shapes)).toEqual(before);
    },
  );

  it("refuses to paint before anything was captured, whatever is selected", async () => {
    const [a] = two();
    helpers.selectShapes([a.id]);
    expect(await refusal(() => tools.applyObjectStyle())).toBe(
      "Capture an object style first.",
    );
  });
});

describe("a selection one object short", () => {
  it.each([
    ["Align", 1, () => tools.alignSelected("left")],
    ["Distribute", 2, () => tools.distributeSelected("vertical")],
    ["Match size", 1, () => tools.matchSelectedSize()],
  ] as [string, number, () => Promise<string>][])(
    "%s refuses %i and moves nothing",
    async (_name, count, run) => {
      const shapes = [
        add(box(10, 20, 30, 20)),
        add(box(60, 50, 20, 30)),
        add(box(110, 100, 40, 10)),
      ].slice(0, count);
      helpers.selectShapes(shapes.map((shape) => shape.id));
      const before = boxesOf(shapes);

      await refusal(run);

      expect(boxesOf(shapes)).toEqual(before);
    },
  );

  it("swaps exactly two and refuses one more", async () => {
    const [a, b] = two();
    const c = add(box(110, 100, 40, 10));
    helpers.selectShapes([a.id, b.id, c.id]);
    expect(await refusal(() => tools.swapSelected())).toBe(
      "Select exactly two objects to swap.",
    );
    helpers.selectShapes([a.id, b.id]);
    expect(await tools.swapSelected()).toBe("Swapped two objects.");
  });
});

describe("the shapes a real slide holds", () => {
  it("aligns an empty layout placeholder like any other object", async () => {
    const placeholder = add({ ...box(200, 90, 120, 40), type: "Placeholder" });
    const [a] = two();
    helpers.selectShapes([a.id, placeholder.id]);

    expect(await tools.alignSelected("left")).toBe("Aligned 2 objects left.");
    expect(placeholder.left).toBe(10);
  });

  it("refuses Select similar on a shape inside a group, keeping the selection", async () => {
    // Clicking into a group selects the child, and the slide's own shape
    // collection cannot see it: matching against the slide would find nothing,
    // clear the selection and still report a success.
    const [a, b] = two();
    presentation.groupShapes([a.id, b.id], presentation.slides[0]!.id);
    helpers.selectShapes([a.id]);

    expect(await refusal(() => tools.selectSimilar())).toBe(
      "Select similar needs an object on the slide, not one inside a group.",
    );
    expect(selectedShapeIds(presentation)).toEqual([a.id]);
  });

  it("still matches the reference itself when nothing else is like it", async () => {
    const only = add(box(0, 0, 100, 50));
    add({ ...box(200, 0, 100, 50), type: "TextBox" });
    helpers.selectShapes([only.id]);

    expect(await tools.selectSimilar()).toBe("Selected 1 similar object.");
    expect(selectedShapeIds(presentation)).toEqual([only.id]);
  });

  it("moves a group's child when the user picked it on purpose", async () => {
    const [a, b] = two();
    const c = add(box(300, 300, 40, 40));
    presentation.groupShapes([a.id, b.id], presentation.slides[0]!.id);
    helpers.selectShapes([a.id, c.id]);

    expect(await tools.alignSelected("top")).toBe("Aligned 2 objects top.");
    expect(a.top).toBe(20);
    expect(c.top).toBe(20);
  });
});

describe("a host that will not take the write", () => {
  it("reports a refused move and takes the next press", async () => {
    const [a, b] = two();
    helpers.selectShapes([a.id, b.id]);
    // The read lands, the write is refused: what a locked shape does.
    helpers.failNextSync(undefined, 1);

    const message = await refusal(() => tools.alignSelected("left"));
    expect(message).toBe("PowerPoint could not complete the request.");

    // The pane is not stuck: the same press works on the next try.
    helpers.selectShapes([a.id, b.id]);
    expect(await tools.alignSelected("left")).toBe("Aligned 2 objects left.");
  });

  it("names the round trip that stopped answering instead of hanging", async () => {
    const [a, b] = two();
    helpers.selectShapes([a.id, b.id]);
    helpers.hangNextSync(1);
    vi.useFakeTimers();

    const message = await refusal(() =>
      settleHungSync(tools.matchSelectedSize()),
    );

    vi.useRealTimers();
    expect(message).toBe(
      "PowerPoint stopped answering while moving the shapes",
    );
  });

  it("refuses every tool below PowerPointApi 1.5 without a round trip", async () => {
    const [a, b] = two();
    helpers.selectShapes([a.id, b.id]);
    helpers.setSupported((_set, version) => version !== "1.5");
    const from = helpers.syncCount();

    for (const [, run] of everyTool()) {
      expect(await refusal(run)).toBe(
        "Object tools need PowerPoint 2021 or Microsoft 365.",
      );
    }
    expect(helpers.syncCount()).toBe(from);
  });
});

describe("Smart Painter against the objects on a slide", () => {
  it("captures a table's own fill and paints a text box that has no line", async () => {
    const table = add({ ...box(0, 0, 200, 80), type: "Table" });
    table.fillColor = "#B27E54";
    table.lineVisible = false;
    const text = add({ ...box(0, 200, 120, 30), type: "TextBox", text: "1,4" });

    helpers.selectShapes([table.id]);
    expect(await tools.captureObjectStyle()).toBe("Object style captured.");
    helpers.selectShapes([text.id]);
    expect(await tools.applyObjectStyle()).toBe("Painted 1 object.");

    expect(text.fillColor).toBe("#B27E54");
    // An invisible captured line paints an invisible line, never a null one.
    expect(text.lineVisible).toBe(false);
    expect(text.lineColor).toBeNull();
    expect(text.text).toBe("1,4");
  });

  it("captures a group's style and paints several targets at once", async () => {
    const [a, b] = two();
    const group = presentation.groupShapes(
      [a.id, b.id],
      presentation.slides[0]!.id,
    );
    group.fillColor = "#282623";
    const targets = [add(box(0, 300, 20, 20)), add(box(40, 300, 20, 20))];

    helpers.selectShapes([group.id]);
    await tools.captureObjectStyle();
    helpers.selectShapes(targets.map((shape) => shape.id));

    expect(await tools.applyObjectStyle()).toBe("Painted 2 objects.");
    expect(targets.map((shape) => shape.fillColor)).toEqual([
      "#282623",
      "#282623",
    ]);
  });

  it("keeps a capture after its source is gone and refuses a picture", async () => {
    const source = add(box(0, 0, 40, 40));
    source.fillColor = "#B27E54";
    const target = add(box(100, 0, 20, 20));
    helpers.selectShapes([source.id]);
    await tools.captureObjectStyle();
    presentation.deleteShape(source.id);

    helpers.selectShapes([target.id]);
    expect(await tools.applyObjectStyle()).toBe("Painted 1 object.");
    expect(target.fillColor).toBe("#B27E54");

    const picture = add({ ...box(200, 0, 40, 40), fillImage: "<png>" });
    helpers.selectShapes([picture.id]);
    expect(await refusal(() => tools.captureObjectStyle())).toBe(
      "Smart Painter currently supports solid or no-fill objects.",
    );
    // The refused capture left the old one in place, rather than half of it.
    helpers.selectShapes([target.id]);
    expect(await tools.applyObjectStyle()).toBe("Painted 1 object.");
  });

  it("paints over a picture fill without asking what was there", async () => {
    const source = add(box(0, 0, 40, 40));
    source.fillColor = "#B27E54";
    const picture = add({ ...box(100, 0, 40, 40), fillImage: "<png>" });
    helpers.selectShapes([source.id]);
    await tools.captureObjectStyle();
    helpers.selectShapes([picture.id]);

    expect(await tools.applyObjectStyle()).toBe("Painted 1 object.");
    expect(picture.fillColor).toBe("#B27E54");
  });
});
