// Small remaining coverage gaps across owned files that do not warrant a
// whole file of their own: LocalCollector's read methods actually thrown,
// relay.support.ts's rejection() on a promise that RESOLVES, and
// links.ts's watchActiveSheet below ExcelApi 1.7 and a failed registration.

import { describe, expect, it, vi } from "vitest";
import { rejection } from "../../src/link/relay.support";
import { LocalCollector } from "../../src/link/local-collector";

const ID = "0123456789abcdef0123456789abcdef";

describe("LocalCollector: every read method throws, not just status/listInbox", () => {
  it("getLink", async () => {
    const collector = new LocalCollector();
    await expect(collector.getLink(ID, "a")).rejects.toThrow(
      /LocalCollector: getLink is not used in Excel/,
    );
  });

  it("getLinkRev", async () => {
    const collector = new LocalCollector();
    await expect(collector.getLinkRev(ID, "a", 1)).rejects.toThrow(
      /LocalCollector: getLinkRev is not used in Excel/,
    );
  });

  it("fetchLinks", async () => {
    const collector = new LocalCollector();
    await expect(collector.fetchLinks([{ id: ID, auth: "a" }])).rejects.toThrow(
      /LocalCollector: fetchLinks is not used in Excel/,
    );
  });

  it("deleteInbox", async () => {
    const collector = new LocalCollector();
    await expect(collector.deleteInbox("ws", "a", ID)).rejects.toThrow(
      /LocalCollector: deleteInbox is not used in Excel/,
    );
  });
});

describe("relay.support.ts rejection(): the resolving half", () => {
  it("answers undefined for a promise that resolves, not just one that rejects", async () => {
    await expect(
      rejection(Promise.resolve("ignored")),
    ).resolves.toBeUndefined();
  });
});

describe("links.ts watchActiveSheet", () => {
  it("does nothing on a host below ExcelApi 1.7: no registration attempted at all", async () => {
    vi.resetModules();
    const { uninstallFakeHost, installFakeHost } = await import("../fakehost");
    uninstallFakeHost();
    const host = installFakeHost({ sheets: ["Model"] });
    host.helpers.setSupported((_set, version) => version !== "1.7");
    const links = await import("../../src/excel/links");
    const handler = vi.fn(async () => undefined);
    // The fake models no onActivated at all yet, so handler would never
    // fire whether or not the guard exists: Excel.run itself is the proof
    // that the below-1.7 guard, not a fake-host limitation, is what stopped
    // registration.
    const run = vi.spyOn(Excel, "run");

    expect(() => links.watchActiveSheet(handler)).not.toThrow();
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(run).not.toHaveBeenCalled();
    expect(handler).not.toHaveBeenCalled();
    uninstallFakeHost();
  });

  it("swallows a registration it cannot make, never an unhandled rejection", async () => {
    vi.resetModules();
    const { uninstallFakeHost, installFakeHost } = await import("../fakehost");
    uninstallFakeHost();
    const host = installFakeHost({ sheets: ["Model"] });
    // The fake models no onActivated event at all yet: a host that DOES
    // support 1.7 but whose registration call itself fails (a real one
    // might refuse the add, or the sync after it) must still be swallowed
    // by watchActiveSheet's own .catch, not surface as an unhandled
    // rejection with nothing to catch it.
    void host;
    const links = await import("../../src/excel/links");
    const rejections: unknown[] = [];
    const onRejection = (reason: unknown): void => {
      rejections.push(reason);
    };
    process.on("unhandledRejection", onRejection);
    try {
      expect(() => links.watchActiveSheet(async () => undefined)).not.toThrow();
      await new Promise((resolve) => setTimeout(resolve, 10));
      expect(rejections).toEqual([]);
    } finally {
      process.off("unhandledRejection", onRejection);
      uninstallFakeHost();
    }
  });
});
