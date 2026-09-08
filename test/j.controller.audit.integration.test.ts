// Controller follow-up to slice J: the two PowerPoint round trips that were
// still bare after the sweep, an object tool's own syncs and the scan's
// group-opening sync, end within the deadline like every other one, naming
// what stopped, and the next action on the same pane works.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createWorkspace } from "../src/link/workspace";
import { SYNC_TIMEOUT_MS } from "../src/ppt/chart-draw";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, seedLink } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";
import type * as ObjectToolsModule from "../src/ppt/object-tools";

enableStrictLoadSemantics();

// Drives the fake clock until the work settles, the way the J suites do.
async function settle<T>(work: Promise<T>): Promise<T> {
  let done = false;
  void work.then(
    () => (done = true),
    () => (done = true),
  );
  for (let step = 0; step < 10 && !done; step += 1) {
    await vi.advanceTimersByTimeAsync(SYNC_TIMEOUT_MS);
  }
  return work;
}

afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

describe("an object tool whose round trip the host swallows", () => {
  let tools: typeof ObjectToolsModule;
  let presentation: FakePresentation;
  let helpers: FakePptHelpers;

  beforeEach(async () => {
    vi.resetModules();
    uninstallFakePpt();
    const host = installFakePpt({ slides: 1 });
    presentation = host.presentation;
    helpers = host.helpers;
    tools = await import("../src/ppt/object-tools");
  });

  it("ends within the deadline and works again afterwards", async () => {
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

    helpers.hangNextSync();
    vi.useFakeTimers();
    await expect(settle(tools.alignSelected("left"))).rejects.toThrow(
      /stopped answering while reading the selection/,
    );
    vi.useRealTimers();

    await tools.alignSelected("left");
    expect(b.left).toBe(10);
  });
});

describe("a grouped deck whose group-opening read the host swallows", () => {
  let links: typeof LinksModule;
  let presentation: FakePresentation;
  let helpers: FakePptHelpers;
  let relay: FakeRelay;

  beforeEach(async () => {
    ({ links, presentation, helpers, relay } = await bootPpt());
  });

  it("ends the scan within the deadline and scans again afterwards", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    const slide = presentation.slides[0]!;
    const picture = slide.shapes[0]!;
    const caption = presentation.addShape(slide, {
      left: 10,
      top: 300,
      width: 200,
      height: 30,
    });
    presentation.groupShapes([picture.id, caption.id], slide.id);

    // The scan's first sync reads the slides; the second opens the group.
    helpers.hangNextSync(1);
    vi.useFakeTimers();
    await expect(settle(links.listLinks(relay))).rejects.toThrow(
      /stopped answering while reading the groups/,
    );
    vi.useRealTimers();

    expect(await links.listLinks(relay)).toHaveLength(1);
  });
});
