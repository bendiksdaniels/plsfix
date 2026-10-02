// The timeout wrapper on its own: a deadline that outlasts a fetch which
// never settles, the abort it hands that fetch, and a timer left clean once a
// normal response arrives. The wiring into RelayClient (the "timeout"
// RelayError kind and its sentence) lives in relay.test.ts.

import { afterEach, describe, expect, it, vi } from "vitest";
import {
  RELAY_TIMEOUT_MS,
  RelayTimeoutError,
  withRelayTimeout,
} from "./relay-timeout";

describe("withRelayTimeout", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it("rejects once the deadline passes, even when fetch never settles", async () => {
    vi.useFakeTimers();
    const fetchImpl = vi.fn(() => new Promise<Response>(() => undefined));
    const settled = withRelayTimeout(
      fetchImpl as unknown as typeof fetch,
      "https://x.test/",
      {},
    ).catch((error: unknown) => error);
    await vi.advanceTimersByTimeAsync(RELAY_TIMEOUT_MS);
    expect(await settled).toBeInstanceOf(RelayTimeoutError);
  });

  it("aborts the signal it handed to fetch once the deadline passes", async () => {
    vi.useFakeTimers();
    let seen: AbortSignal | undefined;
    const fetchImpl = vi.fn((_input: RequestInfo | URL, init?: RequestInit) => {
      seen = init?.signal ?? undefined;
      return new Promise<Response>(() => undefined);
    });
    const settled = withRelayTimeout(
      fetchImpl as unknown as typeof fetch,
      "https://x.test/",
      {},
    ).catch(() => undefined);
    await vi.advanceTimersByTimeAsync(RELAY_TIMEOUT_MS);
    await settled;
    expect(seen?.aborted).toBe(true);
  });

  it("does not fire early: a fetch that settles just under the deadline wins", async () => {
    vi.useFakeTimers();
    const fetchImpl = vi.fn(
      () =>
        new Promise<Response>((resolve) => {
          setTimeout(() => resolve(new Response("ok")), RELAY_TIMEOUT_MS - 1);
        }),
    );
    const call = withRelayTimeout(
      fetchImpl as unknown as typeof fetch,
      "https://x.test/",
      {},
    );
    await vi.advanceTimersByTimeAsync(RELAY_TIMEOUT_MS - 1);
    const response = await call;
    expect(response.status).toBe(200);
  });

  it("clears its timer once fetch resolves, leaving nothing scheduled", async () => {
    vi.useFakeTimers();
    const fetchImpl = vi.fn(async () => new Response("ok"));
    await withRelayTimeout(
      fetchImpl as unknown as typeof fetch,
      "https://x.test/",
      {},
    );
    expect(vi.getTimerCount()).toBe(0);
  });

  it("clears its timer on a rejected fetch too, not only a resolved one", async () => {
    vi.useFakeTimers();
    const fetchImpl = vi.fn(async () => {
      throw new TypeError("offline");
    });
    await expect(
      withRelayTimeout(
        fetchImpl as unknown as typeof fetch,
        "https://x.test/",
        {},
      ),
    ).rejects.toThrow("offline");
    expect(vi.getTimerCount()).toBe(0);
  });
});
