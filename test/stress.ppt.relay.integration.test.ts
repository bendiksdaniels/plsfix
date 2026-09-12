// @vitest-environment jsdom
// Stress pass on what the pane does with a relay that answers badly: 413, 429
// and a batch that dies halfway. Same harness as stress.ppt.pane
// (stress.ppt.support), so these are the real buttons and the real toast.
// Invariant: a bad answer costs one sentence and the row stays actionable -
// never a half-updated deck the pane calls done.

import { afterEach, describe, expect, it, vi } from "vitest";
import { RelayError } from "../src/link/relay";
import {
  bootPane,
  button,
  click,
  copiedDetails,
  deck,
  hasDetails,
  linkRows,
  paneRelay,
  plant,
  pushPlanted,
  settle,
  toastText,
  waiting,
} from "./stress.ppt.support";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import type * as RelayModule from "../src/link/relay";

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
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("a relay that answers badly", () => {
  it("counts a rate-limited update as failed and keeps the reason to copy", async () => {
    await bootPane();
    await plant();
    click("refresh-links");
    await settle();
    for (const call of ["fetchLinks", "getLink"] as const) {
      vi.spyOn(paneRelay(), call).mockRejectedValue(
        new RelayError(
          "server",
          "relay POST /api/links/fetch: 429 too many requests",
          429,
        ),
      );
    }

    click("update-all");
    await settle();

    expect(toastText()).toBe("1 failed");
    expect(hasDetails()).toBe(true);
    expect(button("update-all").disabled).toBe(false);
    // The row is still there to try again with once the minute is over.
    expect(linkRows()).toHaveLength(1);
  });

  it("keeps the rows it reached when the relay dies mid-batch", async () => {
    await bootPane();
    const first = await plant();
    const second = await plant();
    click("refresh-links");
    await settle();
    expect(linkRows()).toHaveLength(2);
    for (const planted of [first, second]) await pushPlanted(planted);
    // The batch answers the first link and defers the second; the GET that
    // would have carried it never reaches the relay.
    const relay = paneRelay();
    const real = relay.fetchLinks.bind(relay);
    vi.spyOn(relay, "fetchLinks").mockImplementation(async (items) => {
      const answer = await real(items.slice(0, 1));
      return {
        items: answer.items,
        omitted: [
          ...answer.omitted,
          ...items.slice(1).map(({ id }) => ({
            id,
            reason: "deferred" as const,
          })),
        ],
      };
    });
    vi.spyOn(relay, "getLink").mockRejectedValue(
      new RelayError("network", "relay GET /api/links: network error"),
    );

    click("update-all");
    await settle();

    expect(toastText()).toBe("1 updated, 1 failed");
    expect(hasDetails()).toBe(true);
    // The row that landed keeps its new picture; the other is untouched.
    const painted = deck().slides[1]!.shapes;
    expect(painted.find((one) => one.id === first.shapeId)!.setImageCalls).toBe(
      1,
    );
    expect(
      painted.find((one) => one.id === second.shapeId)!.setImageCalls,
    ).toBe(0);
    expect(button("update-all").disabled).toBe(false);
  });

  it("says an export is too big to send in words", async () => {
    await bootPane();
    await waiting();
    click("refresh-inbox");
    await settle();
    vi.spyOn(paneRelay(), "getLink").mockRejectedValue(
      new RelayError("tooLarge", "relay GET /api/links: 413", 413),
    );

    click("paste-latest-linked");
    await settle();

    expect(toastText()).toContain("too big to send");
    expect(button("paste-latest-linked").disabled).toBe(false);
  });

  it("reports a relay that cannot be reached and leaves the pane usable", async () => {
    await bootPane();
    await plant();
    click("refresh-links");
    await settle();
    vi.spyOn(paneRelay(), "status").mockRejectedValue(
      new RelayError("network", "relay POST /api/links/status: network error"),
    );

    click("refresh-links");
    await settle();

    expect(toastText()).toBe("relay POST /api/links/status: network error");
    expect(button("refresh-links").disabled).toBe(false);
    expect(linkRows()).toHaveLength(1);
  });

  // The relay's own HTTP sentence is nothing a modeller can act on:
  // src/link/relay-reason.ts gives each answer one of the pane's, the way a
  // 413 has always been reworded.
  it.each([
    [429, "The relay is busy. Try again in a minute."],
    [507, "The relay is full. Ask for space to be cleared."],
    [500, "The relay had a problem. Try again in a minute."],
  ])("says what a %i means instead of showing it", async (status, sentence) => {
    await bootPane();
    await waiting();
    click("refresh-inbox");
    await settle();
    vi.spyOn(paneRelay(), "getLink").mockRejectedValue(
      new RelayError(
        "server",
        `relay GET /api/links: ${String(status)}`,
        status,
      ),
    );

    click("paste-latest-linked");
    await settle();

    expect(toastText()).toContain(sentence);
    expect(toastText()).not.toMatch(/\d{3}|\/api\//);
    expect(button("paste-latest-linked").disabled).toBe(false);
  });

  it("says the same about a row an update could not fetch", async () => {
    await bootPane();
    const planted = await plant();
    click("refresh-links");
    await settle();
    await pushPlanted(planted);
    for (const call of ["fetchLinks", "getLink"] as const) {
      vi.spyOn(paneRelay(), call).mockRejectedValue(
        new RelayError("server", "relay POST /api/links/fetch: 429", 429),
      );
    }

    click("update-all");
    await settle();

    expect(toastText()).toBe("1 failed");
    expect(await copiedDetails()).toContain(
      "Model!B4:F12: The relay is busy. Try again in a minute.",
    );
  });

  // The shape is on the slide before the inbox row is dropped, so a delete the
  // relay refuses must not turn a landed insert into a reported failure: the
  // modeller presses again and lands a second copy.
  it("keeps a landed insert a success when the inbox row cannot be dropped", async () => {
    await bootPane();
    await waiting();
    click("refresh-inbox");
    await settle();
    vi.spyOn(paneRelay(), "deleteInbox").mockRejectedValue(
      new RelayError("server", "relay DELETE /api/inbox: 500", 500),
    );

    click("paste-latest-linked");
    await settle();

    expect(toastText()).toContain("Inserted");
    expect(deck().slides[0]!.shapes).toHaveLength(1);

    // The row is still on the relay, so the Inbox still offers it: pressing
    // again must not leave the slide holding the same picture twice.
    click("paste-latest-linked");
    await settle();
    expect(deck().slides[0]!.shapes).toHaveLength(1);
  });
});
