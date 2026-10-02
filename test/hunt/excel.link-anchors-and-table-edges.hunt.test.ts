// Small, otherwise-uncovered corners of link-anchors.ts and link-table.ts:
// bothRefused's no-code fallback, newEntry actually stamping a project,
// forget() wording a non-Error relay throw, and a table export on a host
// below ExcelApi 1.11 (no separators to read, so no cell is localised).

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { bothRefused, newEntry } from "../../src/excel/link-anchors";
import { renderTable } from "../../src/excel/link-table";
import { newToken } from "../../src/link/crypto";
import { RelayError, type RelayApi } from "../../src/link/relay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "../fakehost";

enableStrictLoadSemantics();

describe("bothRefused / reasonOf", () => {
  it("names the office.js code of the first failure when it has one", () => {
    const first = Object.assign(new Error("GeneralException"), {
      code: "GeneralException",
    });
    const second = new Error("The image failed to render");

    const error = bothRefused(first, "sharp", second, "plain");

    expect(error.message).toBe(
      "sharp: GeneralException; plain: The image failed to render",
    );
  });

  // Neither an office.js error (no `.code`) nor even an Error instance: the
  // fallback must still produce words, not "[object Object]" or a throw.
  it("falls back to String(error) when the first failure has neither a code nor an Error shape", () => {
    const error = bothRefused(
      "network dropped",
      "sharp",
      new Error("plain also failed"),
      "plain",
    );

    expect(error.message).toBe(
      "sharp: network dropped; plain: plain also failed",
    );
  });

  it("reads the second failure's own message when it is a plain Error", () => {
    const error = bothRefused(
      new Error("first"),
      "sharp",
      new Error("second"),
      "plain",
    );
    expect(error.message).toBe("sharp: first; plain: second");
  });
});

describe("newEntry", () => {
  it("stamps the project it is given", () => {
    const entry = newEntry(
      "id",
      "range",
      "PLSFIX_LINK_id",
      "Model!A1",
      "Amasty",
    );
    expect(entry.project).toBe("Amasty");
  });

  it("carries no project field at all when none is given", () => {
    const entry = newEntry("id", "range", "PLSFIX_LINK_id", "Model!A1");
    expect("project" in entry).toBe(false);
  });
});

describe("forget()", () => {
  const entry = {
    id: "a".repeat(32),
    kind: "range" as const,
    anchor: "PLSFIX_LINK_a",
    label: "Model!A1",
    token: newToken(),
    createdAt: "2026-01-01T00:00:00.000Z",
    lastPushedAt: null,
    rev: 1,
  };

  it("reports a RelayError's own message", async () => {
    const { forget } = await import("../../src/excel/link-anchors");
    const relay = {
      deleteLink: () =>
        Promise.reject(new RelayError("network", "relay unreachable")),
    } as unknown as RelayApi;

    await expect(forget(entry, relay)).rejects.toThrow(
      "remove Model!A1: relay unreachable; nothing was removed, try again",
    );
  });

  // The interface's contract is `Promise<void>`, but nothing stops a relay
  // implementation from rejecting with something that is not an Error at
  // all (a raw string, a plain object): the sentence must still read as
  // words, matching reasonOf's own String(error) fallback in link-anchors.ts.
  it("reports a non-Error rejection in words instead of [object Object]", async () => {
    const { forget } = await import("../../src/excel/link-anchors");
    const relay = {
      deleteLink: () => Promise.reject("connection reset"),
    } as unknown as RelayApi;

    await expect(forget(entry, relay)).rejects.toThrow(
      "remove Model!A1: connection reset; nothing was removed, try again",
    );
  });

  it("treats a missing relay copy as done, throwing nothing", async () => {
    const { forget } = await import("../../src/excel/link-anchors");
    const relay = {
      deleteLink: () => Promise.reject(new RelayError("missing", "gone", 404)),
    } as unknown as RelayApi;

    await expect(forget(entry, relay)).resolves.toBeUndefined();
  });
});

describe("renderTable on a host below ExcelApi 1.11", () => {
  let helpers: FakeHelpers;
  afterEach(() => uninstallFakeHost());

  beforeEach(() => {
    vi.resetModules();
    uninstallFakeHost();
    const host = installFakeHost({
      sheets: ["Model"],
      separators: { decimal: ",", thousands: " " },
    });
    helpers = host.helpers;
    // Every ExcelApi check answers true but 1.11 (so this host reads as a
    // real pre-1.11 build, never one refusing everything) - image export
    // (1.9), needed elsewhere in this suite, still answers true too.
    helpers.setSupported((_set, version) => version !== "1.11");
  });

  it("never localises a numeric cell's text: no Application.decimalSeparator to read", async () => {
    helpers.seed("Model!B4", [[8.5]]);

    const render = await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem("Model").getRange("B4");
      return renderTable(context, range);
    });

    // The application's separators are comma/space, but a host that cannot
    // answer for them must leave Office.js's own invariant-format text
    // alone rather than guess.
    expect(render.cells[0]?.[0]?.t).toBe("8.5");
  });
});
