import { describe, expect, it, vi } from "vitest";
import { isRelayError, RelayClient, relayBaseUrl, RelayError } from "./relay";

type Handler = (url: string, init: RequestInit) => Response;
function client(handler: Handler): {
  relay: RelayClient;
  calls: { url: string; init: RequestInit }[];
} {
  const calls: { url: string; init: RequestInit }[] = [];
  const fetchImpl = vi.fn(
    async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      calls.push({ url, init: init ?? {} });
      return handler(url, init ?? {});
    },
  ) as unknown as typeof fetch;
  return {
    relay: new RelayClient("https://x.test/modelis/api/", fetchImpl),
    calls,
  };
}

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
    expect(new Headers(calls[0]!.init.headers).get("x-smt-link-id")).toBe(
      "a".repeat(32),
    );
    expect(calls[0]!.init.body).toEqual(new Uint8Array([5, 6]));
  });
});
