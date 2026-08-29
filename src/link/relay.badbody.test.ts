// One rule, over every route: a 200 whose body is not the shape the route
// promises is a RelayError with a kind, never the SyntaxError a bare
// response.json() would throw and never a value the pane would go on to use.
// The route-by-route calls are in relay.test.ts; both share relay.support.ts.
import { describe, expect, it } from "vitest";
import { isRelayError, type RelayClient } from "./relay";
import { client, rejection } from "./relay.support";

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
      name: "a fetch batch answered with rows instead of the two lists",
      body: new Response(JSON.stringify([{ id, rev: 2, blob: "AQI" }])),
      method: "POST",
      call: (relay) => relay.fetchLinks([{ id, auth: "x", knownRev: 1 }]),
      path: "/modelis/api/links/fetch",
    },
    {
      name: "a fetch batch omitting a link for a reason we do not know",
      body: new Response(
        JSON.stringify({ items: [], omitted: [{ id, reason: "later" }] }),
      ),
      method: "POST",
      call: (relay) => relay.fetchLinks([{ id, auth: "x", knownRev: 1 }]),
      path: "/modelis/api/links/fetch",
    },
    {
      name: "a fetched blob outside the base64url alphabet",
      body: new Response(
        JSON.stringify({
          items: [{ id, rev: 2, blob: "a*b" }],
          omitted: [],
        }),
      ),
      method: "POST",
      call: (relay) => relay.fetchLinks([{ id, auth: "x", knownRev: 1 }]),
      path: "/modelis/api/links/fetch",
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
