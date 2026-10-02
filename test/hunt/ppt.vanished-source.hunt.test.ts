// Hunt pass 1: the six per-link actions pressed three times in a row, and
// again after the modeller deleted the shape a row still names - retagLink
// not wrapping a missing shape like breakLink/refreshLink did. Also pins
// Change source's 429 relay wording and its best-effort inbox cleanup.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { RelayError } from "../../src/link/relay-error";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import type { FakeRelay } from "../fakerelay";
import { fakePng } from "../fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
} from "../fakeppt";
import { bootPpt, memoryStore, pushAgain, seedLink } from "../ppt.support";
import type * as ChangeSourceModule from "../../src/ppt/change-source";
import type * as HostModule from "../../src/ppt/host";
import type * as LinksModule from "../../src/ppt/links";
import type * as RevertModule from "../../src/ppt/revert";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let host: typeof HostModule;
let changeSource: typeof ChangeSourceModule;
let revert: typeof RevertModule;
let presentation: FakePresentation;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
  host = await import("../../src/ppt/host");
  changeSource = await import("../../src/ppt/change-source");
  revert = await import("../../src/ppt/revert");
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  uninstallFakePpt();
});

// No office.js code, no shape id: the same sanity check
// test/stress.ppt.links.integration.test.ts already holds every other action
// to.
function expectSentence(line: string | undefined): void {
  expect(line).toBeDefined();
  expect(line).toContain("Model!B4:F12");
  expect(line).not.toMatch(/ItemNotFound|GeneralException|InvalidArgument/);
}

describe("Change source after the source link vanished", () => {
  it("names the link instead of the host's raw code", async () => {
    const placed = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(placed, ws, relay);
    const row = (await links.listLinks(relay))[0]!;
    const newer = await seedLink(fakePng(400, 200), "Model_v5.xlsx");
    presentation.deleteShape(row.found.shapeId);

    await expect(
      changeSource.changeSource(row, newer, ws, relay),
    ).rejects.toThrow(
      "change source Model!B4:F12: that object is no longer where the list had it.",
    );
  });
});

describe("Go to slide, update, revert, break and change source pressed three times in a row", () => {
  it("update all: idempotent, never stacks a second update", async () => {
    const placed = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(placed, ws, relay);
    await pushAgain(placed, fakePng(400, 200));

    const first = await links.updateLinks(await links.listLinks(relay), relay);
    const second = await links.updateLinks(await links.listLinks(relay), relay);
    const third = await links.updateLinks(await links.listLinks(relay), relay);

    expect(first).toMatchObject({ updated: 1 });
    expect(second).toMatchObject({ updated: 0, current: 1 });
    expect(third).toMatchObject({ updated: 0, current: 1 });
    expect(await links.listLinks(relay)).toMatchObject([{ status: "current" }]);
  });

  it("revert: three presses in a row exhaust the one previous revision and then say so", async () => {
    const placed = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(placed, ws, relay);
    await pushAgain(placed, fakePng(400, 200));
    await links.updateLinks(await links.listLinks(relay), relay);

    const rows = () => links.listLinks(relay);
    const first = await revert.revertLinks(await rows(), relay);
    const second = await revert.revertLinks(await rows(), relay);
    const third = await revert.revertLinks(await rows(), relay);

    expect(first).toMatchObject({ reverted: 1, noPrevious: 0 });
    expect(second).toMatchObject({ reverted: 0, noPrevious: 1 });
    expect(third).toMatchObject({ reverted: 0, noPrevious: 1 });
  });

  it("break: three presses in a row all report the same one link, never a second time removed", async () => {
    const placed = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(placed, ws, relay);
    const row = (await links.listLinks(relay))[0]!;

    await host.breakLink(row.found);
    // A broken link's shape carries no tags any more, so a second and third
    // "Break" on the SAME (now stale) row must still answer cleanly rather
    // than crash on tags that are already gone.
    await expect(host.breakLink(row.found)).resolves.toBeUndefined();
    await expect(host.breakLink(row.found)).resolves.toBeUndefined();
    expect(await links.listLinks(relay)).toHaveLength(0);
  });

  it("go to slide: three presses in a row keep answering the same slide", async () => {
    const placed = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(placed, ws, relay);
    const row = (await links.listLinks(relay))[0]!;

    await expect(host.goToSlide(row.found.slideId)).resolves.toBeUndefined();
    await expect(host.goToSlide(row.found.slideId)).resolves.toBeUndefined();
    await expect(host.goToSlide(row.found.slideId)).resolves.toBeUndefined();
  });

  it("change source: three different presses in a row each consume one inbox item", async () => {
    const placed = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(placed, ws, relay);
    let row = (await links.listLinks(relay))[0]!;

    for (const workbook of [
      "Model_v5.xlsx",
      "Model_v6.xlsx",
      "Model_v7.xlsx",
    ]) {
      const next = await seedLink(fakePng(300, 150), workbook);
      const summary = await changeSource.changeSource(row, next, ws, relay);
      expect(summary).toContain(`-> ${workbook}`);
      row = (await links.listLinks(relay))[0]!;
    }
    expect(row.found.tag.src.workbook).toBe("Model_v7.xlsx");
  });
});

describe("Update this slide after the source link vanished", () => {
  it("reports the deleted shape by name instead of throwing raw ItemNotFound", async () => {
    const placed = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(placed, ws, relay);
    const rows = await links.listLinks(relay);
    await pushAgain(placed, fakePng(400, 200));
    presentation.deleteShape(rows[0]!.found.shapeId);

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expectSentence(summary.failures[0]);
  });
});

describe("Change source when the relay refuses the fetch", () => {
  it("reads a 429 as the same sentence Insert would give, not a raw status line", async () => {
    const placed = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(placed, ws, relay);
    const row = (await links.listLinks(relay))[0]!;
    const newer = await seedLink(fakePng(400, 200), "Model_v5.xlsx");
    relay.getLink = async () => {
      throw new RelayError(
        "server",
        "relay GET /api/links/x: 429 rate limited",
        429,
      );
    };

    await expect(
      changeSource.changeSource(row, newer, ws, relay),
    ).rejects.toThrow(
      "change source Model!B4:F12: The relay is busy. Try again in a minute.",
    );
  });
});

describe("Change source when the inbox cleanup fails after the re-point already landed", () => {
  it("still reports the change instead of the relay's cleanup failure", async () => {
    const placed = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(placed, ws, relay);
    const row = (await links.listLinks(relay))[0]!;
    const newer = await seedLink(fakePng(400, 200), "Model_v5.xlsx");
    relay.deleteInbox = async () => {
      throw new RelayError("server", "relay DELETE /api/inbox/x/y: 503", 503);
    };

    await expect(
      changeSource.changeSource(row, newer, ws, relay),
    ).resolves.toBe("Source changed: Model_v4.xlsx -> Model_v5.xlsx");
    // The picker would otherwise stay open and invite a second Confirm on a
    // link that is already re-pointed (chooser.ts's after() never runs).
    expect((await links.listLinks(relay))[0]!.found.tag.src.workbook).toBe(
      "Model_v5.xlsx",
    );
  });
});
