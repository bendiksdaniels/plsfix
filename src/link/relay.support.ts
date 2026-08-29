// Shared harness for the relay client's suites (relay.test.ts,
// relay.badbody.test.ts): a RelayClient wired to a stubbed fetch that records
// what it was called with, and the rejection of a call as a value. No
// assertions and no Office.js - only the two files' common setup, kept here so
// each suite stays under the file-length limit rather than duplicating it.

import { vi } from "vitest";
import { RelayClient } from "./relay";

export type Handler = (url: string, init: RequestInit) => Response;

export function client(handler: Handler): {
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
export function rejection(call: Promise<unknown>): Promise<unknown> {
  return call.then(
    () => undefined,
    (error: unknown) => error,
  );
}
