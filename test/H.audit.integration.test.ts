// Slice H audit: PowerPoint object tools against groups and non-geometric
// shape types, the Smart Painter line-color null-write fix, the relay
// independence check, and both ribbon command tables' wiring. Extends
// test/ppt.objects.integration.test.ts and test/ppt.commands.integration
// .test.ts with the seeds those files' own fixtures never exercised.

import { readFileSync } from "node:fs";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  selectedShapeIds,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
  type FakeShapeInit,
} from "./fakeppt";
import type * as ObjectToolsModule from "../src/ppt/object-tools";
import type * as CommandsModule from "../src/ppt/commands";

enableStrictLoadSemantics();

let tools: typeof ObjectToolsModule;
let commands: typeof CommandsModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 2 });
  presentation = host.presentation;
  helpers = host.helpers;
  tools = await import("../src/ppt/object-tools");
  commands = await import("../src/ppt/commands");
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

function commandTable(): Record<string, () => Promise<string>> {
  return commands.commandTable({
    notify: () => undefined,
    context: { host: "PowerPoint", version: "test" },
    showTools: async () => undefined,
  });
}

// ---------------------------------------------------------------------------
// Independence: the Tools tab must work with no relay installed at all.
// ---------------------------------------------------------------------------

describe("independence from the relay path", () => {
  it("imports nothing from the relay path in the owned object-tool files", () => {
    for (const file of [
      "../src/ppt/object-math.ts",
      "../src/ppt/object-tools.ts",
      "../src/ppt/commands.ts",
    ]) {
      const text = readFileSync(new URL(file, import.meta.url), "utf8");
      expect(text).not.toMatch(/from ["'][^"']*\/link/);
      expect(text.toLowerCase()).not.toContain("relay");
    }
  });

  it("every object tool works on an unpaired deck with no relay ever installed", async () => {
    // installFakePpt above never constructs a FakeRelay and never touches a
    // workspace/pairing key - the exact "relay unreachable, Links tab never
    // opened" scenario the independence check asks for.
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, box(0, 0, 10, 10));
    const b = presentation.addShape(slide, box(50, 0, 10, 10));
    const c = presentation.addShape(slide, box(100, 0, 10, 10));

    helpers.selectShapes([a.id, b.id, c.id]);
    await expect(tools.alignSelected("left")).resolves.toContain("Aligned");
    helpers.selectShapes([a.id, b.id, c.id]);
    await expect(tools.distributeSelected("horizontal")).resolves.toContain(
      "Distributed",
    );
    helpers.selectShapes([a.id, b.id]);
    await expect(tools.matchSelectedSize()).resolves.toContain("Matched");
    helpers.selectShapes([a.id, b.id]);
    await expect(tools.swapSelected()).resolves.toBe("Swapped two objects.");
    helpers.selectShapes([a.id]);
    await expect(tools.selectSimilar()).resolves.toContain("Selected");
    helpers.selectShapes([a.id]);
    await tools.captureObjectStyle();
    helpers.selectShapes([b.id]);
    await expect(tools.applyObjectStyle()).resolves.toContain("Painted");
  });

  it("every PLSFIX_PPT_* command reaches its handler by name from the table alone", async () => {
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, box(0, 0, 10, 10));
    const b = presentation.addShape(slide, box(50, 0, 10, 10));
    const c = presentation.addShape(slide, box(100, 0, 10, 10));
    const t = commandTable();

    helpers.selectShapes([a.id, b.id, c.id]);
    await expect(t["PLSFIX_PPT_DIST_ACROSS"]!()).resolves.toBe(
      "Distributed 3 objects across.",
    );
    helpers.selectShapes([a.id, b.id, c.id]);
    await expect(t["PLSFIX_PPT_DIST_DOWN"]!()).resolves.toBe(
      "Distributed 3 objects down.",
    );
    helpers.selectShapes([a.id, b.id]);
    await expect(t["PLSFIX_PPT_MATCH"]!()).resolves.toBe(
      "Matched 1 object to the reference size.",
    );
    helpers.selectShapes([a.id]);
    await expect(t["PLSFIX_PPT_CAPTURE"]!()).resolves.toBe(
      "Object style captured.",
    );
    helpers.selectShapes([b.id]);
    await expect(t["PLSFIX_PPT_PAINT"]!()).resolves.toBe("Painted 1 object.");
    helpers.selectShapes([a.id]);
    await expect(t["PLSFIX_PPT_SIMILAR"]!()).resolves.toContain("Selected");
  });

  it("registerCommands notifies success for a real geometry command, not just Object tools", async () => {
    // src/ppt/commands.test.ts only fires PLSFIX_PPT_TOOLS through the full
    // Office.actions.associate wrapper (empty message, no notify call): the
    // "a real command's message !== '' notifies success" branch of that
    // wrapper was never exercised by any existing test.
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, box(0, 0, 30, 20));
    const b = presentation.addShape(slide, box(60, 50, 20, 30));
    helpers.selectShapes([a.id, b.id]);

    type Handler = (event?: { completed: () => void }) => void;
    const associated = new Map<string, Handler>();
    (globalThis as { Office?: unknown }).Office = {
      actions: {
        associate: (id: string, handler: Handler) =>
          associated.set(id, handler),
      },
    };
    const notify = vi.fn();
    commands.registerCommands({
      notify,
      context: { host: "PowerPoint", version: "test" },
      showTools: async () => undefined,
    });
    const completed = vi.fn();
    associated.get("PLSFIX_PPT_SWAP")!({ completed });
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(notify).toHaveBeenCalledWith("Swapped two objects.", "success");
    expect(completed).toHaveBeenCalledOnce();
    expect(presentation.findShape(a.id).shape.left).toBe(60);
    delete (globalThis as { Office?: unknown }).Office;
  });
});

// ---------------------------------------------------------------------------
// Smart Painter: the line.color null-write fix.
// ---------------------------------------------------------------------------

describe("Smart Painter: a visible line with no color of its own", () => {
  it("never writes a null line color onto a target that already had one", async () => {
    const slide = presentation.slides[0]!;
    // A plain, never-recolored shape: lineVisible defaults true, lineColor
    // defaults null (still the theme's), the exact combination captureObjectStyle
    // and paintShape must not confuse with "the line isn't visible".
    const source = presentation.addShape(slide, box(0, 0, 40, 40));
    source.fillColor = "#B27E54";
    const target = presentation.addShape(slide, box(100, 0, 20, 20));
    target.lineColor = "#000000";

    helpers.selectShapes([source.id]);
    await tools.captureObjectStyle();
    helpers.selectShapes([target.id]);
    await expect(tools.applyObjectStyle()).resolves.toBe("Painted 1 object.");

    const { shape } = presentation.findShape(target.id);
    // The captured color was never a real value: the target's own explicit
    // color must survive untouched, not become a literal null.
    expect(shape.lineColor).toBe("#000000");
    expect(shape.lineVisible).toBe(true);
    expect(shape.fillColor).toBe("#B27E54");
  });

  it("paints an explicit line color when the source really had one", async () => {
    const slide = presentation.slides[0]!;
    const source = presentation.addShape(slide, box(0, 0, 40, 40));
    source.fillColor = "#B27E54";
    source.lineColor = "#282623";
    const target = presentation.addShape(slide, box(100, 0, 20, 20));
    target.lineColor = "#000000";

    helpers.selectShapes([source.id]);
    await tools.captureObjectStyle();
    helpers.selectShapes([target.id]);
    await tools.applyObjectStyle();

    expect(presentation.findShape(target.id).shape.lineColor).toBe("#282623");
  });

  it("reports the plural once more than one object is painted", async () => {
    const slide = presentation.slides[0]!;
    const source = presentation.addShape(slide, box(0, 0, 40, 40));
    source.fillColor = "#B27E54";
    const t1 = presentation.addShape(slide, box(100, 0, 20, 20));
    const t2 = presentation.addShape(slide, box(150, 0, 20, 20));

    helpers.selectShapes([source.id]);
    await tools.captureObjectStyle();
    helpers.selectShapes([t1.id, t2.id]);
    await expect(tools.applyObjectStyle()).resolves.toBe("Painted 2 objects.");
  });
});

// ---------------------------------------------------------------------------
// Smart Painter on non-geometric shape types: never a crash, never a "null"
// string, source and target alike.
// ---------------------------------------------------------------------------

describe("Smart Painter on a line, a table and a placeholder", () => {
  it("captures a Line's own stroke: a line has no fill, so it reads NoFill", async () => {
    const slide = presentation.slides[0]!;
    const line = presentation.addShape(slide, {
      ...box(0, 0, 72, 72),
      type: "Line",
    });
    line.lineColor = "#B27E54";
    line.lineWeight = 2;
    const target = presentation.addShape(slide, box(100, 0, 20, 20));
    target.fillColor = "#FFFFFF";

    helpers.selectShapes([line.id]);
    await expect(tools.captureObjectStyle()).resolves.toBe(
      "Object style captured.",
    );
    helpers.selectShapes([target.id]);
    await expect(tools.applyObjectStyle()).resolves.toBe("Painted 1 object.");

    const { shape } = presentation.findShape(target.id);
    expect(shape.fillCleared).toBe(true);
    expect(shape.lineColor).toBe("#B27E54");
    expect(shape.lineWeight).toBe(2);
  });

  it("paints a native table's outer shape without throwing", async () => {
    const slide = presentation.slides[0]!;
    const source = presentation.addShape(slide, box(0, 0, 40, 40));
    source.fillColor = "#B27E54";
    const tableShape = presentation.addShape(slide, {
      ...box(100, 0, 200, 80),
      type: "Table",
    });

    helpers.selectShapes([source.id]);
    await tools.captureObjectStyle();
    helpers.selectShapes([tableShape.id]);
    await expect(tools.applyObjectStyle()).resolves.toBe("Painted 1 object.");
    expect(presentation.findShape(tableShape.id).shape.fillColor).toBe(
      "#B27E54",
    );
  });

  it("paints an empty content placeholder without throwing", async () => {
    const slide = presentation.slides[0]!;
    const source = presentation.addShape(slide, box(0, 0, 40, 40));
    source.fillColor = "#B27E54";
    const placeholder = presentation.addShape(slide, {
      ...box(100, 0, 200, 150),
      type: "Placeholder",
    });

    helpers.selectShapes([source.id]);
    await tools.captureObjectStyle();
    helpers.selectShapes([placeholder.id]);
    await expect(tools.applyObjectStyle()).resolves.toBe("Painted 1 object.");
    expect(presentation.findShape(placeholder.id).shape.fillColor).toBe(
      "#B27E54",
    );
  });
});

// ---------------------------------------------------------------------------
// Groups: a group in the selection is one box, not its children.
// ---------------------------------------------------------------------------

describe("a group in the selection", () => {
  it("getSelectedShapes answers the group itself, sized to its members' bounds", async () => {
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, box(0, 0, 40, 20));
    const b = presentation.addShape(slide, box(60, 10, 20, 30));
    const group = presentation.groupShapes([a.id, b.id], slide.id);

    helpers.selectShapes([group.id]);
    await expect(tools.matchSelectedSize()).rejects.toThrow(
      "Select the reference object first, then at least one target.",
    );
    // The group's own box is the union of its members: left 0, top 0,
    // right 80, bottom 40.
    expect(group.left).toBe(0);
    expect(group.top).toBe(0);
    expect(group.width).toBe(80);
    expect(group.height).toBe(40);
  });

  it("aligns a group as one unit, by its own box, alongside a plain shape", async () => {
    const slide = presentation.slides[0]!;
    const m1 = presentation.addShape(slide, box(20, 0, 40, 20));
    const m2 = presentation.addShape(slide, box(80, 10, 20, 30));
    const group = presentation.groupShapes([m1.id, m2.id], slide.id);
    const plain = presentation.addShape(slide, box(200, 200, 10, 10));

    helpers.selectShapes([group.id, plain.id]);
    const message = await tools.alignSelected("left");
    expect(message).toBe("Aligned 2 objects left.");
    // The group moves as a single box: its own left changes to the
    // selection's minimum (20, the group's own left). Its members' stored
    // coordinates are untouched by this add-in - PowerPoint itself
    // repositions them when it moves a real group.
    expect(presentation.findShape(group.id).shape.left).toBe(20);
    expect(presentation.findShape(plain.id).shape.left).toBe(20);
  });

  it("swaps a group with a plain shape by their own boxes", async () => {
    const slide = presentation.slides[0]!;
    const m1 = presentation.addShape(slide, box(0, 0, 40, 20));
    const m2 = presentation.addShape(slide, box(60, 10, 20, 30));
    const group = presentation.groupShapes([m1.id, m2.id], slide.id);
    const plain = presentation.addShape(slide, box(200, 200, 10, 10));

    helpers.selectShapes([group.id, plain.id]);
    const swapMessage = await tools.swapSelected();
    expect(swapMessage).toBe("Swapped two objects.");
    // Corners trade places; each keeps its own size (the group's box stays
    // 80x40, computed from its members - it is not resized to the plain
    // shape's 10x10).
    expect(presentation.findShape(group.id).shape).toMatchObject({
      left: 200,
      top: 200,
      width: 80,
      height: 40,
    });
    expect(presentation.findShape(plain.id).shape).toMatchObject({
      left: 0,
      top: 0,
      width: 10,
      height: 10,
    });
  });

  it("select similar treats a group as one opaque Group-typed candidate", async () => {
    const slide = presentation.slides[0]!;
    const m1 = presentation.addShape(slide, box(0, 0, 40, 20));
    const m2 = presentation.addShape(slide, box(60, 10, 20, 30));
    const groupA = presentation.groupShapes([m1.id, m2.id], slide.id);
    // A second, same-size group elsewhere on the slide.
    const n1 = presentation.addShape(slide, box(300, 0, 40, 20));
    const n2 = presentation.addShape(slide, box(360, 10, 20, 30));
    const groupB = presentation.groupShapes([n1.id, n2.id], slide.id);
    // A lone top-level shape the same size as one of the group's members:
    // select similar must not reach inside groupA/groupB to "find" it.
    const decoy = presentation.addShape(slide, box(500, 0, 40, 20));

    helpers.selectShapes([groupA.id]);
    const message = await tools.selectSimilar();
    expect(message).toBe("Selected 2 similar objects.");
    expect(selectedShapeIds(presentation).sort()).toEqual(
      [groupA.id, groupB.id].sort(),
    );
    expect(selectedShapeIds(presentation)).not.toContain(decoy.id);
    expect(selectedShapeIds(presentation)).not.toContain(m1.id);
  });

  it("select similar from a shape sub-selected inside a group says so and keeps the selection", async () => {
    // Slide.shapes (and this fake's model) only ever lists top-level shapes;
    // a member of a group is reachable by id (double-clicking into a group
    // and selecting one shape inside it is a real PowerPoint gesture,
    // Presentation.getSelectedShapes can report it), but it is invisible to
    // this top-level scan - it can neither match itself nor be re-selected
    // afterwards, since Slide.setSelectedShapes only accepts top-level ids
    // (learn.microsoft.com/javascript/api/powerpoint/powerpoint.slide,
    // setSelectedShapes: "List of shape IDs to select in the slide").
    const slide = presentation.slides[0]!;
    const member = presentation.addShape(slide, box(0, 0, 40, 20));
    const sibling = presentation.addShape(slide, box(60, 0, 40, 20));
    presentation.groupShapes([member.id, sibling.id], slide.id);
    // No other top-level shape of member's size and type exists on the slide.

    helpers.selectShapes([member.id]);
    // Until the P3 stress pass this answered "Selected 0 similar objects."
    // and set the selection to nothing: a success message for a no-op that
    // took the user's own selection with it. The scan is unchanged; only the
    // answer is, because the source matching nothing can only mean this.
    await expect(tools.selectSimilar()).rejects.toThrow(
      "Select similar needs an object on the slide, not one inside a group.",
    );
    expect(selectedShapeIds(presentation)).toEqual([member.id]);
  });
});
