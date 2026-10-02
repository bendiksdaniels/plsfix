// Hunt pass 2, target 2: the relay answering 408 / 413 / 429 / 500 / 503, a
// timeout, or a garbled body, one call at a time, across insert, Update all,
// revert and Change source. Invariant: the exact sentence the pane shows
// (relay-reason.ts's own line, or the relay's raw message where it has
// none), nothing half-written, the link's content and tag left exactly as
// they were before a failed attempt.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { RelayError, statusKind } from "../../src/link/relay-error";
import { relayReason } from "../../src/link/relay-reason";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import type { FakeRelay } from "../fakerelay";
import { fakePng } from "../fakepng";
import { enableStrictLoadSemantics, uninstallFakePpt } from "../fakeppt";
import { bootPpt, memoryStore, pushAgain, seedLink } from "../ppt.support";
import type * as ChangeSourceModule from "../../src/ppt/change-source";
import type * as LinksModule from "../../src/ppt/links";
import type * as RevertModule from "../../src/ppt/revert";

enableStrictLoadSemantics();

function statusError(status: number, path = "/api/links/x"): RelayError {
  return new RelayError(
    statusKind(status),
    `relay GET ${path}: ${status}`,
    status,
  );
}

const TIMEOUT_ERROR = new RelayError(
  "timeout",
  "The link relay did not answer in time.",
);
const GARBLED_BODY_ERROR = new RelayError(
  "server",
  "relay GET /api/links/x: bad response",
);
// 408 (Request Timeout, server/src/lib.rs's own TimeoutLayer on a stalled
// body): today relay-reason.ts has no line for it (statusKind gives
// "server", but relayReason's own status>=500 check excludes 408), so the
// pane falls back to the raw message below. Its own `sentence` (below) is
// computed from relayReason itself rather than hardcoded, so this row keeps
// passing the day another branch gives 408 a line there.
const ERROR_408 = statusError(408);

// The exact sentence the pane shows for each status: relay-reason.ts's own
// line where it has one, otherwise the relay's raw message - the same
// fallback relay-reason.ts documents for itself. Timeout and garbled body
// pin their raw message directly (relay-reason.ts has no line for either
// kind and never will), so only the 408 row reads relayReason's own answer.
const NAMED_CASES: { label: string; error: RelayError; sentence: string }[] = [
  {
    label: "413 too large",
    error: statusError(413),
    sentence:
      "That export is too big to send. Export a smaller range from Excel.",
  },
  {
    label: "429 rate limited",
    error: statusError(429),
    sentence: "The relay is busy. Try again in a minute.",
  },
  {
    label: "500 server error",
    error: statusError(500),
    sentence: "The relay had a problem. Try again in a minute.",
  },
  {
    label: "503 unavailable",
    error: statusError(503),
    sentence: "The relay had a problem. Try again in a minute.",
  },
  { label: "timeout", error: TIMEOUT_ERROR, sentence: TIMEOUT_ERROR.message },
  {
    label: "garbled body",
    error: GARBLED_BODY_ERROR,
    sentence: GARBLED_BODY_ERROR.message,
  },
  {
    label: "408 request timeout",
    error: ERROR_408,
    sentence: relayReason(ERROR_408) ?? ERROR_408.message,
  },
];

function expectNoRawOfficeCode(message: string): void {
  expect(message).not.toMatch(/InvalidArgument|ItemNotFound|GeneralException/);
}

describe("insert: the relay refuses the initial getLink, one status at a time", () => {
  let links: typeof LinksModule;
  let presentation: Awaited<ReturnType<typeof bootPpt>>["presentation"];
  let relay: FakeRelay;
  let ws: Workspace;

  beforeEach(async () => {
    ({ links, presentation, relay } = await bootPpt());
    ws = await createWorkspace(memoryStore());
  });
  afterEach(() => {
    uninstallFakePpt();
  });

  for (const { label, error, sentence } of NAMED_CASES) {
    it(`${label}: rejects cleanly, no shape ever lands`, async () => {
      const item = await seedLink(fakePng(200, 100));
      relay.getLink = async () => {
        throw error;
      };

      let threw: unknown;
      try {
        await links.insertFromInbox(item, ws, relay);
      } catch (caught) {
        threw = caught;
      }

      expect(threw).toBeInstanceOf(Error);
      const message = (threw as Error).message;
      expectNoRawOfficeCode(message);
      expect(message).toContain(sentence);
      expect(presentation.slides[0]!.shapes).toHaveLength(0);
    });
  }
});

describe("Update all: the relay refuses the fetch, one status at a time", () => {
  let links: typeof LinksModule;
  let presentation: Awaited<ReturnType<typeof bootPpt>>["presentation"];
  let relay: FakeRelay;
  let ws: Workspace;

  beforeEach(async () => {
    ({ links, presentation, relay } = await bootPpt());
    ws = await createWorkspace(memoryStore());
  });
  afterEach(() => {
    uninstallFakePpt();
  });

  for (const { label, error, sentence } of NAMED_CASES) {
    it(`${label}: the row fails, the old picture and tag are untouched`, async () => {
      const item = await seedLink(fakePng(200, 100));
      const placed = await links.insertFromInbox(item, ws, relay);
      await pushAgain(item, fakePng(400, 200));
      const rows = await links.listLinks(relay);
      const before = presentation.slides[0]!.shapes[0]!;
      const oldImage = before.fillImage;
      const oldTag = before.tags.get("PLSFIX_LINK");

      // Both paths a failure could take: the batch fetch, and the per-row
      // GET the batch route falls back to for anything it cannot answer.
      relay.fetchLinks = async () => {
        throw error;
      };
      relay.getLink = async () => {
        throw error;
      };

      const summary = await links.updateLinks(rows, relay);

      expect(summary.updated).toBe(0);
      expect(summary.failed).toBe(1);
      const message = summary.failures[0]!;
      expectNoRawOfficeCode(message);
      expect(message).toContain(sentence);
      const after = presentation.slides[0]!.shapes[0]!;
      expect(after.id).toBe(before.id);
      expect(after.fillImage).toBe(oldImage);
      expect(after.tags.get("PLSFIX_LINK")).toBe(oldTag);
      void placed;
    });
  }
});

describe("Revert: the relay refuses getLinkRev, one status at a time", () => {
  let links: typeof LinksModule;
  let revert: typeof RevertModule;
  let presentation: Awaited<ReturnType<typeof bootPpt>>["presentation"];
  let relay: FakeRelay;
  let ws: Workspace;

  beforeEach(async () => {
    ({ links, presentation, relay } = await bootPpt());
    revert = await import("../../src/ppt/revert");
    ws = await createWorkspace(memoryStore());
  });
  afterEach(() => {
    uninstallFakePpt();
  });

  for (const { label, error, sentence } of NAMED_CASES) {
    it(`${label}: the row fails, the current picture and tag are untouched`, async () => {
      const item = await seedLink(fakePng(200, 100));
      await links.insertFromInbox(item, ws, relay);
      await pushAgain(item, fakePng(400, 200));
      await links.updateLinks(await links.listLinks(relay), relay);
      const before = presentation.slides[0]!.shapes[0]!;
      const currentImage = before.fillImage;
      const currentTag = before.tags.get("PLSFIX_LINK");

      relay.getLinkRev = async () => {
        throw error;
      };
      const summary = await revert.revertLinks(
        await links.listLinks(relay),
        relay,
      );

      expect(summary.reverted).toBe(0);
      expect(summary.failed).toBe(1);
      const message = summary.failures[0]!;
      expectNoRawOfficeCode(message);
      expect(message).toContain(sentence);
      const after = presentation.slides[0]!.shapes[0]!;
      expect(after.fillImage).toBe(currentImage);
      expect(after.tags.get("PLSFIX_LINK")).toBe(currentTag);
    });
  }
});

describe("Change source: the relay refuses the new export's getLink, one status at a time", () => {
  let links: typeof LinksModule;
  let changeSource: typeof ChangeSourceModule;
  let presentation: Awaited<ReturnType<typeof bootPpt>>["presentation"];
  let relay: FakeRelay;
  let ws: Workspace;

  beforeEach(async () => {
    ({ links, presentation, relay } = await bootPpt());
    changeSource = await import("../../src/ppt/change-source");
    ws = await createWorkspace(memoryStore());
  });
  afterEach(() => {
    uninstallFakePpt();
  });

  for (const { label, error, sentence } of NAMED_CASES) {
    it(`${label}: rejects cleanly, the shape still tracks its OLD source`, async () => {
      const item = await seedLink(fakePng(200, 100));
      await links.insertFromInbox(item, ws, relay);
      const row = (await links.listLinks(relay))[0]!;
      const before = presentation.slides[0]!.shapes[0]!;
      const oldTag = before.tags.get("PLSFIX_LINK");
      const newer = await seedLink(fakePng(400, 200), "Model_v5.xlsx");

      const realGetLink = relay.getLink.bind(relay);
      relay.getLink = async (id, auth, knownRev) => {
        if (id === newer.id) throw error;
        return realGetLink(id, auth, knownRev);
      };

      let threw: unknown;
      try {
        await changeSource.changeSource(row, newer, ws, relay);
      } catch (caught) {
        threw = caught;
      }

      expect(threw).toBeInstanceOf(Error);
      const message = (threw as Error).message;
      expectNoRawOfficeCode(message);
      expect(message).toContain(sentence);
      const after = presentation.slides[0]!.shapes[0]!;
      expect(after.tags.get("PLSFIX_LINK")).toBe(oldTag);
    });
  }
});
