// J audit: withSyncDeadline generalised past chart-draw.ts to every other
// PowerPoint.run sync in the pane (host.ts, picture.ts, tables.ts, texts.ts,
// placement.ts). Same shape as ppt.charts.audit.integration.test.ts: a host
// that swallows one round trip must reject within the deadline instead of
// leaving the pane busy for ever, naming what stopped. Strict load semantics
// are on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { type TableCell } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import { SYNC_TIMEOUT_MS } from "../src/ppt/chart-draw";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import {
  bootPpt,
  memoryStore,
  seedLink,
  seedTable,
  seedText,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
const CELLS: TableCell[][] = [[{ t: "Revenue" }, { t: "1 000" }]];
const WIDTHS = [80, 60];

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

function shapes(): FakePptShape[] {
  return presentation.slides[0]!.shapes;
}

// Drives the fake clock until the work under test settles, so a per-sync
// deadline can be proven without the test waiting a real minute. Copied from
// ppt.charts.audit.integration.test.ts rather than shared, because that file
// owns the chart-drawing suite and this one the rest of the pane.
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

// An empty slide's insert is three round trips: which slide is selected, the
// boxes already on it (no placeholders here, so no third read), then the
// insert's own sync. Hanging the third one is the insert itself; hanging the
// first is the placement's own read, with nothing yet on the slide to clean.
const INSERT_SYNC = 2;

describe("a picture insert the host never answers", () => {
  it("rejects instead of hanging, names what stopped, and leaves no shape behind", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(PNG);
    helpers.hangNextSync(INSERT_SYNC);
    vi.useFakeTimers();

    await expect(
      settle(links.insertFromInbox(item, ws, relay)),
    ).rejects.toThrow(
      "PowerPoint stopped answering while inserting the picture",
    );
    expect(shapes()).toHaveLength(0);

    // The next insert on the same pane is not jammed by the one that hung:
    // the item is still on the relay, since the failed attempt never reached
    // the deleteInbox call that follows a successful one.
    vi.useRealTimers();
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(shapes()).toHaveLength(1);
    expect(placed.shapeId).toBe(shapes()[0]!.id);
  });

  it("fails rather than waits when the placement's own read of the slide hangs", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(PNG);
    // The very first round trip of the whole insert - which slide is
    // selected - never answers, so there is nothing on the slide yet either.
    helpers.hangNextSync();
    vi.useFakeTimers();

    await expect(
      settle(links.insertFromInbox(item, ws, relay)),
    ).rejects.toThrow("PowerPoint stopped answering while reading the slide");
    expect(shapes()).toHaveLength(0);
  });
});

describe("a table insert the host never answers", () => {
  it("rejects instead of hanging, names what stopped, and leaves no shape behind", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable(CELLS, WIDTHS);
    helpers.hangNextSync(INSERT_SYNC);
    vi.useFakeTimers();

    await expect(
      settle(links.insertFromInbox(item, ws, relay)),
    ).rejects.toThrow("PowerPoint stopped answering while inserting the table");
    expect(shapes()).toHaveLength(0);

    vi.useRealTimers();
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(shapes()[0]!.type).toBe("Table");
    expect(placed.shapeId).toBe(shapes()[0]!.id);
  });
});

describe("a text insert the host never answers", () => {
  it("rejects instead of hanging, names what stopped, and leaves no shape behind", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedText("EUR 15.7m");
    helpers.hangNextSync(INSERT_SYNC);
    vi.useFakeTimers();

    await expect(
      settle(links.insertFromInbox(item, ws, relay)),
    ).rejects.toThrow("PowerPoint stopped answering while inserting the text");
    expect(shapes()).toHaveLength(0);

    vi.useRealTimers();
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(shapes()[0]!.type).toBe("TextBox");
    expect(placed.shapeId).toBe(shapes()[0]!.id);
  });
});
