// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { describeError, installErrorReporting } from "./report";

const ctx = { host: "Excel", version: "v1.1.000" };
describe("describeError", () => {
  it("keeps the user message short and puts everything else in details", () => {
    const { message, details } = describeError(
      new Error("Select a chart first."),
      ctx,
      "export-chart",
    );
    // The pane's own sentence travels unchanged: only a host error (one
    // carrying an office.js code) is prefixed - the block below has the rule.
    expect(message).toBe("Select a chart first.");
    expect(details).toContain("action: export-chart");
    expect(details).toContain("host: Excel");
    expect(details).toContain("version: v1.1.000");
    expect(details).toContain("Error: Select a chart first.");
  });
  it("names Office.js error codes and debug info when present", () => {
    const error = Object.assign(new Error("x"), {
      code: "ItemNotFound",
      debugInfo: { errorLocation: "Range.load" },
    });
    expect(describeError(error, ctx).details).toContain("code: ItemNotFound");
    expect(describeError(error, ctx).details).toContain("Range.load");
  });
  it("falls back for non-Error throws", () => {
    expect(describeError("nope", ctx).message).toBe(
      "The add-in could not complete that action.",
    );
  });
});

// Read syncs in src/excel are bare host strings by design (test/stress.comps
// and test/stress.comps.pinstripes pin two of them); this is the fix, in the
// one pure half of the pair the pane's real guard wiring cannot reach.
// An office.js error carries a string code; the pane's own Errors never do.
function hostError(message: string, code = "GeneralException"): Error {
  return Object.assign(new Error(message), { code });
}

describe("describeError: staging a bare host message with the running action", () => {
  it("adds the action's label to a bare host string", () => {
    const { message } = describeError(
      hostError("The sync failed."),
      ctx,
      "comps-stats",
    );
    expect(message).toBe("Comps stats: The sync failed.");
  });
  it("leaves the pane's own sentence alone, code or no colon", () => {
    const { message } = describeError(
      new Error("Tick a link in the list first."),
      ctx,
      "update-selected",
    );
    expect(message).toBe("Tick a link in the list first.");
  });
  it("leaves the non-Error fallback message alone", () => {
    const { message } = describeError("nope", ctx, "export-chart");
    expect(message).toBe("The add-in could not complete that action.");
  });
  it("leaves a host message that already carries its own stage alone", () => {
    // "cycle-indent" -> label "Cycle indent", which does not even match this
    // stage's own name: the colon is what exempts it, not the wording.
    const { message } = describeError(
      hostError("Indent cycling: this sheet is protected, nothing was changed"),
      ctx,
      "cycle-indent",
    );
    expect(message).toBe(
      "Indent cycling: this sheet is protected, nothing was changed",
    );
  });
  it("leaves a host message opening with the action's own label alone", () => {
    const { message } = describeError(
      hostError("Pinstripes need at least two rows in the selection."),
      ctx,
      "pinstripes-rows",
    );
    expect(message).toBe("Pinstripes need at least two rows in the selection.");
  });
  it("leaves the message unchanged when no action is given", () => {
    const { message } = describeError(hostError("The sync failed."), ctx);
    expect(message).toBe("The sync failed.");
  });
});
describe("installErrorReporting", () => {
  it("forwards window errors and unhandled rejections", () => {
    const notify = vi.fn();
    const listeners: Record<string, (event: unknown) => void> = {};
    const fakeWindow = {
      addEventListener: (type: string, fn: (event: unknown) => void) => {
        listeners[type] = fn;
      },
    };
    installErrorReporting(ctx, notify, fakeWindow as unknown as Window);
    listeners["error"]!({ error: new Error("late") });
    listeners["unhandledrejection"]!({ reason: new Error("async") });
    expect(notify).toHaveBeenCalledTimes(2);
    expect(notify.mock.calls[1]![0]).toBe("async");
  });
});
