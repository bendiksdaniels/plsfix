import { describe, expect, it } from "vitest";
import { RelayError } from "./relay-error";
import {
  relayReason,
  RELAY_BUSY,
  RELAY_FAULT,
  RELAY_FULL,
  TOO_LARGE,
} from "./relay-reason";

describe("relayReason", () => {
  it.each([
    [
      "a 413",
      new RelayError("tooLarge", "413 payload too large", 413),
      TOO_LARGE,
    ],
    [
      "a 429",
      new RelayError("server", "429 too many requests", 429),
      RELAY_BUSY,
    ],
    [
      "a 507",
      new RelayError("server", "507 insufficient storage", 507),
      RELAY_FULL,
    ],
    ["a 500", new RelayError("server", "500 internal", 500), RELAY_FAULT],
    ["a 502", new RelayError("server", "502 bad gateway", 502), RELAY_FAULT],
  ] as [string, RelayError, string][])(
    "answers %s in the pane's words",
    (_name, error, sentence) => {
      expect(relayReason(error)).toBe(sentence);
    },
  );

  it.each([
    ["a network error", new RelayError("network", "network error")],
    ["a 404", new RelayError("missing", "not found", 404)],
    ["a 403", new RelayError("auth", "forbidden", 403)],
    ["a 400", new RelayError("server", "400 too many items", 400)],
  ] as [string, RelayError][])(
    "leaves %s to travel as the relay said it",
    (_name, error) => {
      expect(relayReason(error)).toBeUndefined();
    },
  );

  it("says nothing about an error that is not the relay's", () => {
    expect(relayReason(new Error("ItemNotFound"))).toBeUndefined();
    expect(relayReason("exploded")).toBeUndefined();
    expect(relayReason(undefined)).toBeUndefined();
  });
});
