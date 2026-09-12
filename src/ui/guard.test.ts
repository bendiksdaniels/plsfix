import { describe, expect, it, vi } from "vitest";
import { makeGuard } from "./guard";

describe("makeGuard", () => {
  it("busy on, run, after, decorate, notify, finally, busy off", async () => {
    const calls: string[] = [];
    const guard = makeGuard({
      setBusy: (b) => calls.push(`busy:${b}`),
      notify: (m, k) => calls.push(`notify:${k ?? "success"}:${m}`),
      describe: (e) => ({ message: String(e), details: "" }),
      after: async () => {
        calls.push("after");
      },
      decorate: (m) => `${m}!`,
      finally: () => calls.push("finally"),
    });
    await guard(async () => "Done");
    expect(calls).toEqual([
      "busy:true",
      "after",
      "notify:success:Done!",
      "finally",
      "busy:false",
    ]);
  });
  it("routes a failure to notify(error) with details and still clears busy", async () => {
    const notify = vi.fn();
    const guard = makeGuard({
      setBusy: () => undefined,
      notify,
      describe: (e, action) => ({
        message: (e as Error).message,
        details: `action=${action}`,
      }),
    });
    await guard(async () => {
      throw new Error("boom");
    }, "export");
    expect(notify).toHaveBeenCalledWith("boom", "error", "action=export");
  });

  it("runs one action at a time once busyMessage is set", async () => {
    const notify = vi.fn();
    const guard = makeGuard({
      setBusy: () => undefined,
      notify,
      describe: (e) => ({ message: String(e), details: "" }),
      busyMessage: "Wait for the last action to finish.",
    });
    let release = (): void => undefined;
    const held = new Promise<void>((done) => {
      release = () => done();
    });
    const second = vi.fn(async () => "second");

    const first = guard(async () => {
      await held;
      return "first";
    });
    await guard(second, "update-all");

    expect(second).not.toHaveBeenCalled();
    expect(notify).toHaveBeenCalledWith(
      "Wait for the last action to finish.",
      "error",
    );
    release();
    await first;
    // The latch clears with the flow: the next press runs.
    await guard(second);
    expect(second).toHaveBeenCalledTimes(1);
  });

  it("leaves the latch off for a guard with no busyMessage", async () => {
    const running: string[] = [];
    const guard = makeGuard({
      setBusy: () => undefined,
      notify: () => undefined,
      describe: (e) => ({ message: String(e), details: "" }),
    });
    let release = (): void => undefined;
    const held = new Promise<void>((done) => {
      release = () => done();
    });

    const first = guard(async () => {
      await held;
      return "first";
    });
    await guard(async () => {
      running.push("second");
      return "second";
    });

    // The Excel pane's guard is unchanged: the second action still runs.
    expect(running).toEqual(["second"]);
    release();
    await first;
  });
});
