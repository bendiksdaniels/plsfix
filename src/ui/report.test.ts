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
