import { beforeEach, describe, expect, it, vi } from "vitest";
import type { InboxItem } from "../link/model";
import type * as QueueModule from "./inbox-queue";

const SRC = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

function item(id: string): InboxItem {
  return {
    id,
    token: `token-${id}`,
    kind: "range",
    label: `Model!B4:F12 (${id})`,
    src: SRC,
    createdAt: new Date().toISOString(),
  };
}

function row(
  id: string,
  createdAt: number,
): { item: InboxItem; createdAt: number } {
  return { item: item(id), createdAt };
}

// The pasted set lives for the pane's lifetime, so every test gets a new one.
let queue: typeof QueueModule;

beforeEach(async () => {
  vi.resetModules();
  queue = await import("./inbox-queue");
});

describe("queueOrder", () => {
  it("puts the newest export first whatever order the relay sent", () => {
    const ordered = queue.queueOrder([
      row("older", 1_000),
      row("newest", 3_000),
      row("middle", 2_000),
    ]);
    expect(ordered.map((one) => one.id)).toEqual(["newest", "middle", "older"]);
    expect(queue.latestInboxItem(ordered)?.id).toBe("newest");
  });

  it("drops an export this pane already pasted, however often it is asked", () => {
    queue.rememberPasted("pasted");

    const ordered = queue.queueOrder([
      row("pasted", 3_000),
      row("fresh", 1_000),
    ]);

    expect(ordered.map((one) => one.id)).toEqual(["fresh"]);
    expect(queue.latestInboxItem(ordered)?.id).toBe("fresh");
  });

  it("answers nothing when every waiting row is one this pane pasted", () => {
    queue.rememberPasted("a");
    expect(queue.queueOrder([row("a", 1_000)])).toEqual([]);
    expect(queue.latestInboxItem([])).toBeNull();
  });
});
