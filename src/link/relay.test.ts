// The relay client route by route: what each call puts on the wire (method,
// bearer, headers, body), what it reads back off it, and how every non-success
// outcome reaches the pane as a typed RelayError carrying the reason the relay
// named. The bodies a 200 can still be unusable with live in
// relay.badbody.test.ts; both suites share relay.support.ts.
import { afterEach, describe, expect, it, vi } from "vitest";
import { isRelayError, RelayClient, relayBaseUrl, RelayError } from "./relay";
import { RELAY_TIMEOUT_MS } from "./relay-timeout";
import { client, rejection } from "./relay.support";

describe("relayBaseUrl", () => {
  it("sits beside the pane page", () => {
    expect(
      relayBaseUrl("https://dbautomatizacijas.com/modelis/pptpane.html").href,
    ).toBe("https://dbautomatizacijas.com/modelis/api/");
    expect(relayBaseUrl("https://localhost:3000/taskpane.html").href).toBe(
      "https://localhost:3000/api/",
    );
  });
});

describe("isRelayError", () => {
  it("recognises a relay error the class identity no longer matches", () => {
    const real = new RelayError("missing", "not found", 404);
    expect(isRelayError(real)).toBe(true);
    // What a second module graph hands back: same shape, another constructor.
    const foreign = Object.assign(new Error("forbidden"), {
      name: "RelayError",
      kind: "auth",
    });
    expect(foreign instanceof RelayError).toBe(false);
    expect(isRelayError(foreign)).toBe(true);
    expect(isRelayError(foreign) && foreign.kind).toBe("auth");
  });

  it("rejects anything that is not one", () => {
    expect(isRelayError(new Error("plain"))).toBe(false);
    expect(
      isRelayError(Object.assign(new Error("x"), { name: "RelayError" })),
    ).toBe(false);
    expect(
      isRelayError(
        Object.assign(new Error("x"), { name: "RelayError", kind: "teapot" }),
      ),
    ).toBe(false);
    expect(isRelayError({ name: "RelayError", kind: "auth" })).toBe(false);
    expect(isRelayError(null)).toBe(false);
    expect(isRelayError("RelayError")).toBe(false);
  });
});

describe("RelayClient", () => {
  it("PUTs a link with bearer auth and returns the rev", async () => {
    const { relay, calls } = client(
      () => new Response(JSON.stringify({ rev: 4 }), { status: 200 }),
    );
    const result = await relay.putLink(
      "a".repeat(32),
      "AUTH",
      new Uint8Array([1, 2]),
    );
    expect(result).toEqual({ rev: 4 });
    expect(calls[0]!.url).toBe(
      `https://x.test/modelis/api/links/${"a".repeat(32)}`,
    );
    expect(calls[0]!.init.method).toBe("PUT");
    expect(new Headers(calls[0]!.init.headers).get("authorization")).toBe(
      "Bearer AUTH",
    );
    expect(new Headers(calls[0]!.init.headers).get("content-type")).toBe(
      "application/octet-stream",
    );
  });
  it("GET honours If-None-Match and maps 304", async () => {
    const { relay, calls } = client(() => new Response(null, { status: 304 }));
    expect(await relay.getLink("a".repeat(32), "AUTH", 4)).toBe("unchanged");
    expect(new Headers(calls[0]!.init.headers).get("if-none-match")).toBe(
      '"4"',
    );
  });
  it("GET returns bytes and the ETag rev", async () => {
    const { relay } = client(
      () =>
        new Response(new Uint8Array([9, 9]), {
          status: 200,
          headers: { ETag: '"7"' },
        }),
    );
    const result = await relay.getLink("a".repeat(32), "AUTH");
    expect(result).toEqual({ rev: 7, blob: new Uint8Array([9, 9]) });
  });
  it("GET accepts the weak ETag an intermediary may hand back", async () => {
    const { relay } = client(
      () =>
        new Response(new Uint8Array([9]), {
          status: 200,
          headers: { ETag: 'W/"12"' },
        }),
    );
    expect(await relay.getLink("a".repeat(32), "AUTH")).toEqual({
      rev: 12,
      blob: new Uint8Array([9]),
    });
  });
  it("GET ?rev names one revision and never sends If-None-Match", async () => {
    const { relay, calls } = client(
      () =>
        new Response(new Uint8Array([4, 2]), {
          status: 200,
          headers: { ETag: '"1"' },
        }),
    );
    const result = await relay.getLinkRev("a".repeat(32), "AUTH", 1);
    expect(result).toEqual({ rev: 1, blob: new Uint8Array([4, 2]) });
    expect(calls[0]!.url).toBe(
      `https://x.test/modelis/api/links/${"a".repeat(32)}?rev=1`,
    );
    // A named revision is the one thing a 304 must never answer: the pane is
    // asking for a picture it does not hold.
    expect(new Headers(calls[0]!.init.headers).get("if-none-match")).toBe(null);
  });
  it("GET ?rev maps a dropped revision to missing", async () => {
    const { relay } = client(() => new Response("", { status: 404 }));
    const error = await rejection(relay.getLinkRev("a".repeat(32), "AUTH", 1));
    expect(isRelayError(error) && error.kind).toBe("missing");
  });
  it("GET ?rev refuses an ETag that is not the revision asked for", async () => {
    const { relay } = client(
      () => new Response(new Uint8Array([9]), { status: 200, headers: {} }),
    );
    const error = await rejection(relay.getLinkRev("a".repeat(32), "AUTH", 2));
    expect(isRelayError(error) && error.kind).toBe("server");
    expect(String(error)).toContain("bad ETag");
  });
  it("GET refuses a missing or unparsable ETag instead of reading rev 0", async () => {
    const headerSets: Record<string, string>[] = [
      {},
      { ETag: "" },
      { ETag: '"x"' },
      { ETag: "7" },
    ];
    for (const headers of headerSets) {
      const { relay } = client(
        () => new Response(new Uint8Array([9]), { status: 200, headers }),
      );
      const error = await rejection(relay.getLink("a".repeat(32), "AUTH"));
      expect(isRelayError(error) && error.kind).toBe("server");
      expect(String(error)).toContain("bad ETag");
    }
  });
  it("maps statuses to error kinds", async () => {
    for (const [status, kind] of [
      [403, "auth"],
      [404, "missing"],
      [413, "tooLarge"],
      [500, "server"],
    ] as const) {
      const { relay } = client(() => new Response("", { status }));
      await expect(relay.getLink("a".repeat(32), "AUTH")).rejects.toMatchObject(
        { kind, status } satisfies Partial<RelayError>,
      );
    }
    const failing = new RelayClient("https://x.test/api/", (async () => {
      throw new TypeError("offline");
    }) as unknown as typeof fetch);
    await expect(failing.getLink("a".repeat(32), "AUTH")).rejects.toMatchObject(
      { kind: "network" },
    );
  });
  // Every refusal the relay makes names itself in the body; without it a 400
  // reaches the pane as a bare status code, and "your batch is too long" is
  // indistinguishable from "your id is malformed".
  it("puts the relay's reason for refusing into the message", async () => {
    const { relay } = client(
      () =>
        new Response(JSON.stringify({ error: "too many items" }), {
          status: 400,
        }),
    );
    const error = await rejection(relay.status([{ id: "a", auth: "x" }]));
    expect(isRelayError(error) && error.kind).toBe("server");
    expect(String(error)).toBe(
      "RelayError: relay POST /modelis/api/links/status: 400 too many items",
    );
  });

  it("carries the reason on every route and kind, not just status", async () => {
    for (const [status, kind, reason] of [
      [400, "server", "bad rev"],
      [403, "auth", "another key owns this"],
      [404, "missing", "not found"],
      [500, "server", "store error"],
    ] as const) {
      const { relay } = client(
        () => new Response(JSON.stringify({ error: reason }), { status }),
      );
      const error = await rejection(relay.getLinkRev("a".repeat(32), "K", 3));
      expect(isRelayError(error) && error.kind).toBe(kind);
      expect(String(error)).toBe(
        `RelayError: relay GET /modelis/api/links/${"a".repeat(32)}: ${String(status)} ${reason}`,
      );
    }
  });

  // A refusal with no body at all, one that is not JSON (an Access
  // interstitial), and one shaped differently: the status is still the answer,
  // so the message keeps it and adds nothing rather than failing to be built.
  it("falls back to the bare status when the body names nothing", async () => {
    const bodies = [
      new Response("", { status: 400 }),
      new Response("<html>denied</html>", { status: 400 }),
      new Response(JSON.stringify({ message: "nope" }), { status: 400 }),
      new Response(JSON.stringify({ error: 17 }), { status: 400 }),
    ];
    for (const body of bodies) {
      const { relay } = client(() => body.clone());
      const error = await rejection(relay.status([{ id: "a", auth: "x" }]));
      expect(String(error)).toBe(
        "RelayError: relay POST /modelis/api/links/status: 400",
      );
    }
  });

  it("batches status queries", async () => {
    const { relay, calls } = client(
      () =>
        new Response(
          JSON.stringify([
            { id: "a", rev: 2, pushedAt: 10 },
            { id: "b", rev: null, pushedAt: null, error: "auth" },
          ]),
        ),
    );
    const rows = await relay.status([
      { id: "a", auth: "x" },
      { id: "b", auth: "y" },
    ]);
    expect(rows[1]).toEqual({
      id: "b",
      rev: null,
      pushedAt: null,
      error: "auth",
    });
    expect(JSON.parse(String(calls[0]!.init.body))).toEqual([
      { id: "a", auth: "x" },
      { id: "b", auth: "y" },
    ]);
  });
  it("batches a fetch and decodes the blobs it answers with", async () => {
    const { relay, calls } = client(
      () =>
        new Response(
          JSON.stringify({
            items: [{ id: "a", rev: 3, blob: "AQI" }],
            omitted: [
              { id: "b", reason: "unchanged" },
              { id: "c", reason: "deferred" },
            ],
          }),
        ),
    );
    const result = await relay.fetchLinks([
      { id: "a", auth: "x", knownRev: 2 },
      { id: "b", auth: "y", knownRev: 1 },
      { id: "c", auth: "z" },
    ]);
    expect(result.items).toEqual([
      { id: "a", rev: 3, blob: new Uint8Array([1, 2]) },
    ]);
    expect(result.omitted).toEqual([
      { id: "b", reason: "unchanged" },
      { id: "c", reason: "deferred" },
    ]);
    expect(calls[0]!.url).toBe("https://x.test/modelis/api/links/fetch");
    expect(calls[0]!.init.method).toBe("POST");
    // No bearer: a batch carries a key per item, like the status poll.
    expect(new Headers(calls[0]!.init.headers).get("authorization")).toBe(null);
    expect(JSON.parse(String(calls[0]!.init.body))).toEqual([
      { id: "a", auth: "x", knownRev: 2 },
      { id: "b", auth: "y", knownRev: 1 },
      { id: "c", auth: "z" },
    ]);
  });
  it("lists inbox rows with decoded blobs", async () => {
    const { relay } = client(
      () =>
        new Response(
          JSON.stringify([{ id: "a".repeat(32), createdAt: 5, blob: "AQI" }]),
        ),
    );
    const rows = await relay.listInbox("WS", "AUTH");
    expect(rows[0]!.blob).toEqual(new Uint8Array([1, 2]));
  });
  // Controller ruling (not in the brief): postInbox carries the link id in a
  // header, since the relay reads it off the wire rather than the sealed body.
  it("POSTs an inbox item with the link id header and bearer auth", async () => {
    const { relay, calls } = client(() => new Response(null, { status: 200 }));
    await relay.postInbox("WS", "AUTH", "a".repeat(32), new Uint8Array([5, 6]));
    expect(calls[0]!.url).toBe("https://x.test/modelis/api/inbox/WS");
    expect(calls[0]!.init.method).toBe("POST");
    expect(new Headers(calls[0]!.init.headers).get("authorization")).toBe(
      "Bearer AUTH",
    );
    expect(new Headers(calls[0]!.init.headers).get("content-type")).toBe(
      "application/octet-stream",
    );
    expect(new Headers(calls[0]!.init.headers).get("x-plsfix-link-id")).toBe(
      "a".repeat(32),
    );
    expect(calls[0]!.init.body).toEqual(new Uint8Array([5, 6]));
  });
});

// The wrapper's own mechanics (the abort, the cleared timer) live in
// relay-timeout.test.ts; here only the wiring - a call past RELAY_TIMEOUT_MS
// reaches the pane as the one sentence every consumer's default error path
// already shows verbatim, on both panes, with no per-route rewording needed.
describe("RelayClient timeout", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it("maps a call that never answers to a timeout RelayError", async () => {
    vi.useFakeTimers();
    const fetchImpl = vi.fn(() => new Promise<Response>(() => undefined));
    const relay = new RelayClient(
      "https://x.test/modelis/api/",
      fetchImpl as unknown as typeof fetch,
    );
    const settled = rejection(relay.getLink("a".repeat(32), "AUTH"));
    await vi.advanceTimersByTimeAsync(RELAY_TIMEOUT_MS);
    const error = await settled;
    expect(isRelayError(error) && error.kind).toBe("timeout");
    expect(String(error)).toBe(
      "RelayError: The link relay did not answer in time.",
    );
  });

  it("still reports a genuine network failure as kind network, not timeout", async () => {
    const failing = new RelayClient("https://x.test/api/", (async () => {
      throw new TypeError("offline");
    }) as unknown as typeof fetch);
    const error = await rejection(failing.getLink("a".repeat(32), "AUTH"));
    expect(isRelayError(error) && error.kind).toBe("network");
  });
});
