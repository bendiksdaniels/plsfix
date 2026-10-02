// POST /api/links/touch from the client's side: what it puts on the wire, how
// it cuts a workbook past the route's ceiling into batches and sums the answer,
// and how a refusal or a mis-shaped 200 reaches the caller. The fake relay's
// own touch rules are checked here too, because a fake that drifts from
// server/src/store.rs is a production bug no other suite can see.
import { describe, expect, it } from "vitest";
import { FakeRelay } from "../../test/fakerelay";
import { isRelayError, MAX_TOUCH_ITEMS } from "./relay";
import { client, rejection } from "./relay.support";

const ID = "a".repeat(32);

function items(count: number): { id: string; auth: string }[] {
  return Array.from({ length: count }, (_, index) => ({
    id: String(index).padStart(32, "0"),
    auth: "AUTH",
  }));
}

function counted(touched: number): Response {
  return new Response(JSON.stringify({ touched }), { status: 200 });
}

describe("RelayClient.touchLinks", () => {
  it("posts the pairs and answers with the relay's count", async () => {
    const { relay, calls } = client(() => counted(2));
    expect(await relay.touchLinks(items(2))).toBe(2);
    expect(calls).toHaveLength(1);
    expect(calls[0]!.url).toBe("https://x.test/modelis/api/links/touch");
    expect(calls[0]!.init.method).toBe("POST");
    expect(new Headers(calls[0]!.init.headers).get("content-type")).toBe(
      "application/json",
    );
    // No bearer: this batch carries a key per item, like status and fetch.
    expect(new Headers(calls[0]!.init.headers).get("authorization")).toBe(null);
    expect(JSON.parse(String(calls[0]!.init.body))).toEqual(items(2));
  });

  it("cuts a workbook past the ceiling into batches and sums them", async () => {
    const { relay, calls } = client((_url, init) =>
      counted((JSON.parse(String(init.body)) as unknown[]).length),
    );
    const workbook = items(MAX_TOUCH_ITEMS + 30);
    expect(await relay.touchLinks(workbook)).toBe(workbook.length);
    expect(calls).toHaveLength(2);
    expect(JSON.parse(String(calls[0]!.init.body))).toHaveLength(
      MAX_TOUCH_ITEMS,
    );
    expect(JSON.parse(String(calls[1]!.init.body))).toEqual(
      workbook.slice(MAX_TOUCH_ITEMS),
    );
  });

  it("sends nothing at all for an empty workbook", async () => {
    const { relay, calls } = client(() => counted(0));
    expect(await relay.touchLinks([])).toBe(0);
    expect(calls).toHaveLength(0);
  });

  it("maps a refusal to a kind and keeps the relay's reason", async () => {
    for (const [status, kind, reason] of [
      [400, "server", "too many items"],
      [429, "server", "too many requests"],
      [500, "server", "store error"],
    ] as const) {
      const { relay } = client(
        () => new Response(JSON.stringify({ error: reason }), { status }),
      );
      const error = await rejection(relay.touchLinks(items(1)));
      expect(isRelayError(error) && error.kind).toBe(kind);
      expect(String(error)).toBe(
        `RelayError: relay POST /modelis/api/links/touch: ${String(status)} ${reason}`,
      );
    }
  });

  it("treats a 200 without a count as a bad response", async () => {
    for (const body of ["{}", '{"touched":"2"}', "not json"]) {
      const { relay } = client(() => new Response(body, { status: 200 }));
      const error = await rejection(relay.touchLinks(items(1)));
      expect(isRelayError(error) && error.kind).toBe("server");
      expect(String(error)).toBe(
        "RelayError: relay POST /modelis/api/links/touch: bad response",
      );
    }
  });
});

describe("FakeRelay.touchLinks", () => {
  it("counts only the links the key owns, like the store does", async () => {
    const relay = new FakeRelay();
    await relay.putLink(ID, "MINE", new Uint8Array([1]));
    expect(
      await relay.touchLinks([
        { id: ID, auth: "MINE" },
        { id: ID, auth: "THEIRS" },
        { id: "b".repeat(32), auth: "MINE" },
      ]),
    ).toBe(1);
  });

  it("refuses a batch past the route's ceiling", async () => {
    const relay = new FakeRelay();
    const error = await rejection(relay.touchLinks(items(MAX_TOUCH_ITEMS + 1)));
    expect(isRelayError(error) && error.status).toBe(400);
    expect(String(error)).toBe(
      "RelayError: relay POST /api/links/touch: 400 too many items",
    );
  });
});
