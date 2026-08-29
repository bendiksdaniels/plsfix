// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { copyText } from "./clipboard";

describe("copyText", () => {
  it("copies through the async Clipboard API when the host allows it", async () => {
    const writeText = vi.fn(() => Promise.resolve());
    Object.assign(navigator, { clipboard: { writeText } });

    await copyText("PLSFIX_KEY_abc123");

    expect(writeText).toHaveBeenCalledWith("PLSFIX_KEY_abc123");
  });

  it("falls back to a hidden textarea when navigator.clipboard is undefined", async () => {
    Object.assign(navigator, { clipboard: undefined });
    const execCommand = vi.fn(() => true);
    Object.assign(document, { execCommand });

    await copyText("PLSFIX_KEY_abc123");

    expect(execCommand).toHaveBeenCalledWith("copy");
    expect(document.querySelectorAll("textarea")).toHaveLength(0);
  });
});
