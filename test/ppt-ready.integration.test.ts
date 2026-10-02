// @vitest-environment jsdom
// test/ppt-ready.ts held to its promise under slow key derivation: the pane's
// boot and a refresh both await WebCrypto on Node's thread pool, which the
// full suite slows down, so the helpers must wait for the work, not a count.
import { afterEach, describe, expect, it, vi } from "vitest";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import type * as RelayModule from "../src/link/relay";
import { RelayError } from "../src/link/relay";
import {
  bootPane,
  click,
  paneRelay,
  plant,
  settle,
  toastText,
} from "./stress.ppt.support";

enableStrictLoadSemantics();

vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return paneRelay();
    },
  };
});

afterEach(() => {
  vi.restoreAllMocks();
  uninstallFakePpt();
});

// A thread pool under the full suite's load: every key import waits first.
function slowKeyImports(ms: number): void {
  const subtle = globalThis.crypto.subtle;
  const real = subtle.importKey.bind(subtle) as unknown as (
    ...args: unknown[]
  ) => Promise<CryptoKey>;
  vi.spyOn(subtle, "importKey").mockImplementation(
    (...args: unknown[]) =>
      new Promise<CryptoKey>((resolve, reject) => {
        setTimeout(() => {
          real(...args).then(resolve, reject);
        }, ms);
      }),
  );
}

describe("the PowerPoint pane helpers under slow key derivation", () => {
  it("wait for a slow boot and a slow refresh, not for a count of turns", async () => {
    slowKeyImports(25);
    const spies: { mockRestore(): void }[] = [];
    await bootPane(true, async () => {
      spies.push(
        vi
          .spyOn(paneRelay(), "status")
          .mockRejectedValue(
            new RelayError(
              "network",
              "relay POST /api/links/status: network error",
            ),
          ),
        vi
          .spyOn(paneRelay(), "listInbox")
          .mockRejectedValue(
            new RelayError("network", "relay GET /api/inbox: network error"),
          ),
      );
      await plant();
    });
    expect(toastText()).toBe("relay GET /api/inbox: network error");

    // The relay answers again; the slow key imports stay for the refresh.
    for (const spy of spies) spy.mockRestore();
    click("refresh-links");
    await settle();
    expect(toastText()).toBe("1 linked object.");
  });
});
