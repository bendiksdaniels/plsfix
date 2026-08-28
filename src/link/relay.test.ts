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

// What a call rejected with, so a test can assert on the kind and the message
// of the same object rather than matching a shape twice.
function rejection(call: Promise<unknown>): Promise<unknown> {
  return call.then(
    () => undefined,
    (error: unknown) => error,
  );
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

// A 200 is not a promise that the body is what the route documents: a captive
// portal, an Access interstitial or the dev proxy answering with index.html
// all arrive as 200. Every one of these must reach the pane as a RelayError
// with a kind, never as the SyntaxError a bare response.json() would throw.
describe("RelayClient rejects a 200 whose body is not the promised shape", () => {
  const id = "a".repeat(32);
  const garbage = () => new Response("<!doctype html><html></html>");

  const cases: {
    name: string;
    body: Response;
    method: string;
    call: (relay: RelayClient) => Promise<unknown>;
    path: string;
  }[] = [
    {
      name: "PUT answered with HTML",
      body: garbage(),
      method: "PUT",
      call: (relay) => relay.putLink(id, "AUTH", new Uint8Array([1])),
      path: `/modelis/api/links/${id}`,
    },
    {
      name: "PUT answered with JSON that has no rev",
      body: new Response(JSON.stringify({ ok: true })),
      method: "PUT",
      call: (relay) => relay.putLink(id, "AUTH", new Uint8Array([1])),
      path: `/modelis/api/links/${id}`,
    },
    {
      name: "status answered with HTML",
      body: garbage(),
      method: "POST",
      call: (relay) => relay.status([{ id, auth: "x" }]),
      path: "/modelis/api/links/status",
    },
    {
      name: "status answered with an object instead of rows",
      body: new Response(JSON.stringify({ id, rev: 2 })),
      method: "POST",
      call: (relay) => relay.status([{ id, auth: "x" }]),
      path: "/modelis/api/links/status",
    },
    {
      name: "a status row whose rev is a string",
      body: new Response(JSON.stringify([{ id, rev: "2", pushedAt: 1 }])),
      method: "POST",
      call: (relay) => relay.status([{ id, auth: "x" }]),
      path: "/modelis/api/links/status",
    },
    {
      name: "inbox answered with HTML",
      body: garbage(),
      method: "GET",
      call: (relay) => relay.listInbox("WS", "AUTH"),
      path: "/modelis/api/inbox/WS",
    },
    {
      name: "an inbox row with a null blob",
      body: new Response(JSON.stringify([{ id, createdAt: 5, blob: null }])),
      method: "GET",
      call: (relay) => relay.listInbox("WS", "AUTH"),
      path: "/modelis/api/inbox/WS",
    },
    {
      name: "an inbox blob outside the base64url alphabet",
      body: new Response(JSON.stringify([{ id, createdAt: 5, blob: "a*b" }])),
      method: "GET",
      call: (relay) => relay.listInbox("WS", "AUTH"),
      path: "/modelis/api/inbox/WS",
    },
  ];

  for (const { name, body, method, call, path } of cases) {
    it(`fails as kind "server" for ${name}`, async () => {
      const { relay } = client(() => body.clone());
      const error = await rejection(call(relay));
      expect(isRelayError(error) && error.kind).toBe("server");
      expect(String(error)).toBe(
        `RelayError: relay ${method} ${path}: bad response`,
      );
    });
  }
});
