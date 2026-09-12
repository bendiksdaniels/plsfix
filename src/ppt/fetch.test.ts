// Direct proof of fetchUpdates' relay-batching contract, against the real
// crypto and the in-memory relay every ppt.*.integration.test.ts suite
// already trusts (test/fakerelay.ts): falling back to a per-row GET when a
// batch cannot be answered whole, never asking for more than MAX_FETCH_ITEMS
// links at once, and following up once a batch's blobs cross FETCH_BLOB_CAP.
import { describe, expect, it } from "vitest";
import { FakeRelay } from "../../test/fakerelay";
import { deriveLinkKeys, newToken, seal } from "../link/crypto";
import {
  encodePayload,
  type LinkKind,
  type PicturePayload,
  type Source,
  type TextPayload,
} from "../link/model";
import {
  FETCH_BLOB_CAP,
  MAX_FETCH_ITEMS,
  RelayError,
  type FetchQuery,
  type FetchResult,
} from "../link/relay";
import { fetchUpdates } from "./fetch";
import type { FoundLink } from "./host";
import type { LinkRow } from "./links";

const SRC: Source = {
  workbook: "Model.xlsx",
  sheet: "Model",
  ref: "B2",
  anchor: "PLSFIX_LINK_00000000",
};
const NOW = "2026-09-12T00:00:00.000Z";

function textPayload(text: string): TextPayload {
  return { v: 1, kind: "text", text, src: SRC, pushedAt: NOW, hash: "h" };
}

function picturePayload(pngChars: number): PicturePayload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 10,
    height: 10,
    png: "A".repeat(pngChars),
    src: SRC,
    pushedAt: NOW,
    hash: "h",
  };
}

function found(
  id: string,
  token: string,
  rev: number,
  kind: LinkKind = "text",
): FoundLink {
  return {
    slideId: "1",
    shapeId: "s1",
    slideIndex: 0,
    tag: { v: 1, id, kind, rev, src: SRC, pushedAt: NOW },
    token,
    type: "Picture",
    left: 0,
    top: 0,
    width: 100,
    height: 100,
  };
}

function row(
  id: string,
  token: string,
  rev: number,
  kind: LinkKind = "text",
): LinkRow {
  return {
    found: found(id, token, rev, kind),
    status: "current",
    relayRev: null,
    pushedAt: null,
  };
}

// Stores a payload behind a fresh id/token pair the way a real push would.
async function seedLink(
  relay: FakeRelay,
  id: string,
  token: string,
  payload: TextPayload | PicturePayload,
): Promise<void> {
  const keys = await deriveLinkKeys(token);
  await relay.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
  );
}

describe("fetchUpdates", () => {
  it("counts a link the relay already has current as up to date, without opening it", async () => {
    const relay = new FakeRelay();
    const token = newToken();
    const id = "link-current";
    await seedLink(relay, id, token, textPayload("v1"));
    // The row already knows the revision putLink just created (1): the relay
    // must answer "unchanged" and fetchUpdates must not try to decode it.
    const outcome = await fetchUpdates([row(id, token, 1)], relay);

    expect(outcome).toEqual({
      batch: [],
      current: 1,
      missing: 0,
      wrongKey: 0,
      failures: [],
    });
  });
});

describe("when the relay cannot answer a whole batch", () => {
  it("falls back to a per-row GET when the refusal is not a network failure", async () => {
    class ServerRefusingRelay extends FakeRelay {
      // An old relay before the batch route existed: the server itself
      // answers, just not with a body this client understands.
      override async fetchLinks(): Promise<FetchResult> {
        throw new RelayError("server", "relay POST /api/links/fetch: 404", 404);
      }
    }
    const relay = new ServerRefusingRelay();
    const token = newToken();
    const id = "link-a";
    await seedLink(relay, id, token, textPayload("fallback text"));
    // rev 0 is older than the rev putLink actually created (1), so the
    // fallback GET must return the real blob rather than "unchanged".
    const outcome = await fetchUpdates([row(id, token, 0)], relay);

    expect(outcome.current).toBe(0);
    expect(outcome.failures).toEqual([]);
    expect(outcome.batch).toHaveLength(1);
    expect(outcome.batch[0]!.payload).toEqual(textPayload("fallback text"));
    expect(outcome.batch[0]!.rev).toBe(1);
  });

  it("fails every row outright when the relay cannot be reached, without trying a GET", async () => {
    class UnreachableRelay extends FakeRelay {
      getLinkCalls = 0;
      override async fetchLinks(): Promise<FetchResult> {
        throw new RelayError(
          "network",
          "relay POST /api/links/fetch: network error",
        );
      }
      override async getLink(
        id: string,
        auth: string,
        knownRev?: number,
      ): ReturnType<FakeRelay["getLink"]> {
        this.getLinkCalls += 1;
        return super.getLink(id, auth, knownRev);
      }
    }
    const relay = new UnreachableRelay();
    const token = newToken();
    const id = "link-b";
    await seedLink(relay, id, token, textPayload("never opened"));
    const outcome = await fetchUpdates([row(id, token, 0)], relay);

    expect(outcome.batch).toEqual([]);
    expect(outcome.current).toBe(0);
    expect(outcome.failures).toHaveLength(1);
    expect(outcome.failures[0]!.found.tag.id).toBe(id);
    expect(relay.getLinkCalls).toBe(0);
  });
});

describe("chunking a large batch", () => {
  it("never asks the relay for more than MAX_FETCH_ITEMS links in one call", async () => {
    class CountingRelay extends FakeRelay {
      batchSizes: number[] = [];
      override async fetchLinks(items: FetchQuery[]): Promise<FetchResult> {
        this.batchSizes.push(items.length);
        return super.fetchLinks(items);
      }
    }
    const relay = new CountingRelay();
    // One shared token for every link: fetchUpdates derives a token's key
    // once (keysByToken) and this test is about the chunk boundary, not key
    // handling, so sharing it is a simplification, not a shortcut around it.
    const token = newToken();
    const keys = await deriveLinkKeys(token);
    const total = MAX_FETCH_ITEMS + 1;
    const rows: LinkRow[] = [];
    for (let i = 0; i < total; i += 1) {
      const id = `link-${String(i).padStart(4, "0")}`;
      const payload = textPayload(`row ${String(i)}`);
      await relay.putLink(
        id,
        keys.auth,
        await seal(keys.enc, id, encodePayload(payload)),
      );
      rows.push(row(id, token, 0));
    }

    const outcome = await fetchUpdates(rows, relay);

    expect(outcome.batch).toHaveLength(total);
    expect(outcome.failures).toEqual([]);
    expect(relay.batchSizes.length).toBeGreaterThanOrEqual(2);
    expect(relay.batchSizes.every((size) => size <= MAX_FETCH_ITEMS)).toBe(
      true,
    );
    expect(relay.batchSizes.reduce((sum, size) => sum + size, 0)).toBe(total);
    // Decoding survives the boundary itself: the first item of the overflow
    // chunk must come back exactly as it was pushed.
    const overflowId = `link-${String(MAX_FETCH_ITEMS).padStart(4, "0")}`;
    const overflow = outcome.batch.find(
      (entry) => entry.found.tag.id === overflowId,
    );
    expect(overflow?.payload).toEqual(
      textPayload(`row ${String(MAX_FETCH_ITEMS)}`),
    );
  });
});

describe("the relay's response byte cap", () => {
  it("defers whatever does not fit in one answer and re-fetches it on its own", async () => {
    class CountingGetRelay extends FakeRelay {
      getLinkCalls = 0;
      override async getLink(
        id: string,
        auth: string,
        knownRev?: number,
      ): ReturnType<FakeRelay["getLink"]> {
        this.getLinkCalls += 1;
        return super.getLink(id, auth, knownRev);
      }
    }
    const relay = new CountingGetRelay();
    const token = newToken();
    const keys = await deriveLinkKeys(token);
    // Two blobs a little over 60% of the cap each: the first alone fits, but
    // together they cross FETCH_BLOB_CAP, so the second must be deferred -
    // never sliced mid-batch, per fetchLinks' own "the rest is deferred with
    // it" rule (test/fakerelay.ts, mirroring server/src/fetch.rs).
    const chunkChars = Math.floor(FETCH_BLOB_CAP * 0.6);
    const idA = "link-big-a";
    const idB = "link-big-b";
    const payloadA = picturePayload(chunkChars);
    const payloadB = picturePayload(chunkChars);
    await relay.putLink(
      idA,
      keys.auth,
      await seal(keys.enc, idA, encodePayload(payloadA)),
    );
    await relay.putLink(
      idB,
      keys.auth,
      await seal(keys.enc, idB, encodePayload(payloadB)),
    );
    const rows = [row(idA, token, 0, "range"), row(idB, token, 0, "range")];

    const outcome = await fetchUpdates(rows, relay);

    expect(outcome.failures).toEqual([]);
    expect(outcome.batch).toHaveLength(2);
    const gotA = outcome.batch.find((entry) => entry.found.tag.id === idA)!;
    const gotB = outcome.batch.find((entry) => entry.found.tag.id === idB)!;
    expect(gotA.payload).toEqual(payloadA);
    expect(gotB.payload).toEqual(payloadB);
    // Exactly the overflowing link needed the fallback GET - not zero, not both.
    expect(relay.getLinkCalls).toBe(1);
  });
});
