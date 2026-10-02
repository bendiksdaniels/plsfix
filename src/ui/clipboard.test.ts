// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { copyText } from "./clipboard";

describe("copyText", () => {
  it("copies through the async Clipboard API when the host allows it", async () => {
    const writeText = vi.fn(() => Promise.resolve());
    Object.assign(navigator, { clipboard: { writeText } });

    const copied = await copyText("PLSFIX_KEY_abc123");

    expect(writeText).toHaveBeenCalledWith("PLSFIX_KEY_abc123");
    expect(copied).toBe(true);
  });

  it("falls back to a hidden textarea when navigator.clipboard is undefined, and reports execCommand's true", async () => {
    Object.assign(navigator, { clipboard: undefined });
    const execCommand = vi.fn(() => true);
    Object.assign(document, { execCommand });

    const copied = await copyText("PLSFIX_KEY_abc123");

    expect(execCommand).toHaveBeenCalledWith("copy");
    expect(document.querySelectorAll("textarea")).toHaveLength(0);
    expect(copied).toBe(true);
  });

  it("reports execCommand's false instead of claiming the copy landed", async () => {
    Object.assign(navigator, { clipboard: undefined });
    const execCommand = vi.fn(() => false);
    Object.assign(document, { execCommand });

    const copied = await copyText("PLSFIX_KEY_abc123");

    expect(execCommand).toHaveBeenCalledWith("copy");
    expect(copied).toBe(false);
  });
});
