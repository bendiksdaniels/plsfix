import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { hostReady } from "./host-ready";

const sleep = (ms: number) =>
  new Promise<void>((resolve) => setTimeout(resolve, ms));

const timing = { headStartMs: 4000, probeEveryMs: 1000, giveUpAfterMs: 10000 };

describe("hostReady", () => {
  beforeEach(() => vi.useFakeTimers());
  afterEach(() => vi.useRealTimers());

  it("takes Office.onReady as is and never probes", async () => {
    const probe = vi.fn(() => Promise.resolve<string | null>("Excel"));
    const result = hostReady({
      onReady: Promise.resolve("Excel"),
      probe,
      sleep,
      ...timing,
    });
    await vi.advanceTimersByTimeAsync(20000);
    await expect(result).resolves.toEqual({ host: "Excel", degraded: false });
    expect(probe).not.toHaveBeenCalled();
  });

  it("accepts a host that answers the probe after the head start", async () => {
    const answers: (string | null)[] = [null, "Excel"];
    const probe = vi.fn(() => Promise.resolve(answers.shift() ?? null));
    const result = hostReady({
      onReady: new Promise<string>(() => undefined),
      probe,
      sleep,
      ...timing,
    });
    await vi.advanceTimersByTimeAsync(3999);
    expect(probe).not.toHaveBeenCalled();
    await vi.advanceTimersByTimeAsync(2000);
    await expect(result).resolves.toEqual({ host: "Excel", degraded: true });
    expect(probe).toHaveBeenCalledTimes(2);
  });

  it("keeps waiting for onReady once the probes give up", async () => {
    let ready: (host: string) => void = () => undefined;
    const probe = vi.fn(() => Promise.resolve<string | null>(null));
    const result = hostReady({
      onReady: new Promise<string>((resolve) => (ready = resolve)),
      probe,
      sleep,
      ...timing,
    });
    let done = false;
    void result.then(() => (done = true));
    await vi.advanceTimersByTimeAsync(30000);
    expect(done).toBe(false);
    expect(probe.mock.calls.length).toBeGreaterThan(0);
    const calls = probe.mock.calls.length;
    await vi.advanceTimersByTimeAsync(30000);
    expect(probe).toHaveBeenCalledTimes(calls);
    ready("Excel");
    await expect(result).resolves.toEqual({ host: "Excel", degraded: false });
  });

  it("treats a throwing probe as no answer", async () => {
    const probe = vi
      .fn<() => Promise<string | null>>()
      .mockRejectedValueOnce(new Error("not yet"))
      .mockResolvedValue("Excel");
    const result = hostReady({
      onReady: new Promise<string>(() => undefined),
      probe,
      sleep,
      ...timing,
    });
    await vi.advanceTimersByTimeAsync(6000);
    await expect(result).resolves.toEqual({ host: "Excel", degraded: true });
  });
});
