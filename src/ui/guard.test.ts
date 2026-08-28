import { describe, expect, it, vi } from "vitest";
import { makeGuard } from "./guard";

describe("makeGuard", () => {
  it("busy on, run, after, decorate, notify, finally, busy off", async () => {
    const calls: string[] = [];
    const guard = makeGuard({
      setBusy: (b) => calls.push(`busy:${b}`),
      notify: (m, k) => calls.push(`notify:${k ?? "success"}:${m}`),
      describe: (e) => ({ message: String(e), details: "" }),
      after: async () => { calls.push("after"); },
      decorate: (m) => `${m}!`,
      finally: () => calls.push("finally"),
    });
    await guard(async () => "Done");
    expect(calls).toEqual(["busy:true", "after", "notify:success:Done!", "finally", "busy:false"]);
  });
  it("routes a failure to notify(error) with details and still clears busy", async () => {
    const notify = vi.fn();
    const guard = makeGuard({
      setBusy: () => undefined,
      notify,
      describe: (e, action) => ({ message: (e as Error).message, details: `action=${action}` }),
    });
    await guard(async () => { throw new Error("boom"); }, "export");
    expect(notify).toHaveBeenCalledWith("boom", "error", "action=export");
  });
});
