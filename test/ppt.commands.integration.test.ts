// The PowerPoint ribbon commands against the strict fake host: a ribbon
// button runs the same object tool the pane's button does, on the same
// selection, with the same result line. Split from ppt.objects.integration
// .test.ts, which sits at the file cap.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import type * as CommandsModule from "../src/ppt/commands";

enableStrictLoadSemantics();

let commands: typeof CommandsModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  presentation = host.presentation;
  helpers = host.helpers;
  commands = await import("../src/ppt/commands");
});
afterEach(() => {
  uninstallFakePpt();
});

function table(): Record<string, () => Promise<string>> {
  return commands.commandTable({
    notify: () => undefined,
    context: { host: "PowerPoint", version: "test" },
    showTools: async () => undefined,
  });
}

describe("ribbon commands run the pane's tools", () => {
  it("Swap exchanges the two selected shapes' positions", async () => {
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, {
      left: 10,
      top: 20,
      width: 30,
      height: 20,
    });
    const b = presentation.addShape(slide, {
      left: 60,
      top: 50,
      width: 20,
      height: 30,
    });
    helpers.selectShapes([a.id, b.id]);

    expect(await table()["PLSFIX_PPT_SWAP"]!()).toBe("Swapped two objects.");
    expect(presentation.findShape(a.id).shape).toMatchObject({
      left: 60,
      top: 50,
    });
    expect(presentation.findShape(b.id).shape).toMatchObject({
      left: 10,
      top: 20,
    });
  });

  it("Align left lines the selection up on its leftmost edge", async () => {
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, {
      left: 10,
      top: 20,
      width: 30,
      height: 20,
    });
    const b = presentation.addShape(slide, {
      left: 60,
      top: 50,
      width: 20,
      height: 30,
    });
    helpers.selectShapes([a.id, b.id]);

    expect(await table()["PLSFIX_PPT_ALIGN_LEFT"]!()).toBe(
      "Aligned 2 objects left.",
    );
    expect(presentation.findShape(b.id).shape.left).toBe(10);
  });

  it("refuses like the pane when the selection is too small", async () => {
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, {
      left: 10,
      top: 20,
      width: 30,
      height: 20,
    });
    helpers.selectShapes([a.id]);
    await expect(table()["PLSFIX_PPT_DIST_DOWN"]!()).rejects.toThrow(
      "Select at least three objects to distribute.",
    );
  });
});
