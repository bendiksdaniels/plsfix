// A hard deadline on one relay fetch: a manual AbortController plus
// setTimeout, since AbortSignal.timeout is missing on the older Office
// webviews this pane still has to run inside. relay.ts is the only caller.
// Invariant: the timer is cleared in `finally` on every path - a settled
// fetch, a timeout, or any other rejection - so nothing keeps running after.

export const RELAY_TIMEOUT_MS = 20_000;

// What a call past RELAY_TIMEOUT_MS rejects with. fetch is raced against a
// timer rather than only awaited, so a host that never settles at all still
// loses to the clock; the same timer aborts the signal fetch was given, so a
// real request's socket is freed rather than left running unread.
export class RelayTimeoutError extends Error {
  constructor(ms: number) {
    super(`relay call did not answer within ${String(ms)}ms`);
    this.name = "RelayTimeoutError";
  }
}

export async function withRelayTimeout(
  fetchImpl: typeof fetch,
  input: string,
  init: RequestInit,
  ms: number = RELAY_TIMEOUT_MS,
): Promise<Response> {
  const controller = new AbortController();
  let timer: ReturnType<typeof setTimeout> | undefined;
  const deadline = new Promise<never>((_resolve, reject) => {
    timer = setTimeout(() => {
      controller.abort();
      reject(new RelayTimeoutError(ms));
    }, ms);
  });
  try {
    return await Promise.race([
      fetchImpl(input, { ...init, signal: controller.signal }),
      deadline,
    ]);
  } finally {
    clearTimeout(timer);
  }
}
