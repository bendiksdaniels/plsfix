import { describe, expect, it } from "vitest";
import { relativeTime } from "./time";

describe("relativeTime", () => {
  it("reads never for a timestamp that was never set", () => {
    expect(relativeTime(null)).toBe("never");
  });

  it("buckets by minute, hour and day", () => {
    expect(relativeTime(1000 - 30, 1000)).toBe("just now");
    expect(relativeTime(1000 - 120, 1000)).toBe("2 min ago");
    expect(relativeTime(1000 - 5 * 3600, 1000)).toBe("5 h ago");
    expect(relativeTime(1000 - 3 * 86400, 1000)).toBe("3 days ago");
  });

  it("never goes negative when the clock is ahead of the server", () => {
    expect(relativeTime(1000 + 500, 1000)).toBe("just now");
  });
});
