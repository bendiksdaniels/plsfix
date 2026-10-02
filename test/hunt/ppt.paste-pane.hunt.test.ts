// @vitest-environment jsdom
// Hunt pass 1: the paste box end to end (main.ts's pasteFromExcel through
// bundleFromPaste and bundle.ts's decoder): malformed/foreign HTML, an empty
// or newer bundle, relay mode with no store open, the "forgets between
// sessions" hint, and a partly-failed paste's sentence reaching Copy details.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import type * as RelayModule from "../../src/link/relay";
import { bundleHtml, encodeBundle } from "../../src/link/bundle";
import { deriveLinkKeys, newToken, seal } from "../../src/link/crypto";
import { localWorkspace } from "../../src/link/local";
import { LocalCollector } from "../../src/link/local-collector";
import { TRANSPORT_STORAGE_KEY } from "../../src/link/transport-setting";
import {
  encodeInboxItem,
  encodePayload,
  encodeTag,
  newLinkId,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
  type Payload,
} from "../../src/link/model";
import {
  createWorkspace,
  WORKSPACE_STORAGE_KEY,
  type Workspace,
} from "../../src/link/workspace";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePresentation,
} from "../fakeppt";
import { FakeRelay } from "../fakerelay";
import { fakePng } from "../fakepng";
import { settlePpt, trackPptBoot } from "../ppt-ready";
import { copiedDetails } from "../stress.ppt.support";

enableStrictLoadSemantics();

let relay: FakeRelay;
vi.mock("../../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return relay;
    },
  };
});

const SRC = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

function newId(): string {
  return newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
}

function picturePayload(png: string): Payload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png,
    src: SRC,
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
}

// Seals the link with `sealWith` (defaults to its own token's, the honest
// case) but always addresses the inbox item by the link's real `id`/`token`,
// so a caller can plant a shape that expects to open it and watch that fail.
// `existing` lets a caller plant the shape FIRST (its tag needs the id/token
// before this ever runs) and export into that same identity afterwards -
// pushedAt order matters: assertFresh refuses an "update" whose payload is
// not newer than the tag already on the shape, so planting must come first,
// exactly as test/clickthrough.ppt.integration.test.ts's own plantLink does.
async function exportLinkLocally(
  collector: LocalCollector,
  sealWith?: Awaited<ReturnType<typeof deriveLinkKeys>>,
  existing?: { id: string; token: string },
): Promise<{ id: string; token: string }> {
  const id = existing?.id ?? newId();
  const token = existing?.token ?? newToken();
  const keys = sealWith ?? (await deriveLinkKeys(token));
  await collector.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(picturePayload(fakePng(200, 100)))),
    0,
  );
  const localWs = await localWorkspace();
  const item: InboxItem = {
    id,
    token,
    kind: "range",
    label: "Model!B4:F12",
    src: SRC,
    createdAt: new Date().toISOString(),
  };
  await collector.postInbox(
    localWs.id,
    localWs.auth,
    id,
    await seal(localWs.enc, localWs.id, encodeInboxItem(item)),
  );
  return { id, token };
}

// A shape already on slide 1, tagged at rev 1: the deck this link's paste is
// about to find and try to repaint.
function plantLink(
  presentation: FakePresentation,
  id: string,
  token: string,
): void {
  const shape = presentation.addShape(presentation.slides[1]!, {
    type: "GeometricShape",
    fillImage: fakePng(800, 400),
    left: 100,
    top: 80,
    width: 400,
    height: 200,
  });
  const tag: LinkTag = {
    v: 1,
    id,
    kind: "range",
    rev: 1,
    src: SRC,
    pushedAt: new Date().toISOString(),
  };
  shape.tags.set(TAG_LINK, encodeTag(tag));
  shape.tags.set(TAG_KEY, token);
}

interface BootOptions {
  local?: boolean;
}

let presentation: FakePresentation;

async function bootPpt(options: BootOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  relay = new FakeRelay();
  const storage = new Map<string, string>();
  if (options.local) {
    storage.set(TRANSPORT_STORAGE_KEY, "local");
  } else {
    const workspace: Workspace = await createWorkspace({
      get: async () => null,
      set: async () => undefined,
      remove: async () => undefined,
    });
    storage.set(WORKSPACE_STORAGE_KEY, workspace.exportKey);
    storage.set(TRANSPORT_STORAGE_KEY, "relay");
  }
  ({ presentation } = installFakePpt({ slides: 3, storage }));
  const booted = trackPptBoot();
  await import("../../src/ppt/main");
  await booted;
  await settlePpt();
}

afterEach(() => {
  uninstallFakePpt();
});

function pasteBox(): HTMLTextAreaElement {
  return document.getElementById("paste-links") as HTMLTextAreaElement;
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function pasteEvent(html: string, plain: string): Event {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => (type === "text/html" ? html : plain) },
  });
  return event;
}

describe("a malformed or foreign paste never carries a raw code or JSON to the toast", () => {
  it("a data-plsfix-links attribute holding garbage, not base64url, reads as 'not a copy'", async () => {
    await bootPpt({ local: true });

    pasteBox().dispatchEvent(
      pasteEvent(
        '<span data-plsfix-links="not-real-base64!!!">junk</span>',
        "pls,fix links for PowerPoint: paste them in the pls,fix pane, Inbox tab.",
      ),
    );
    await settlePpt();

    expect(toastText()).toBe("That is not a pls,fix copy from Excel.");
  });

  it("a data-plsfix-links attribute holding valid base64url of non-JSON reads as 'not a copy'", async () => {
    await bootPpt({ local: true });
    const garbage = btoa("not json at all").replaceAll("=", "");

    pasteBox().dispatchEvent(
      pasteEvent(`<span data-plsfix-links="${garbage}">junk</span>`, ""),
    );
    await settlePpt();

    expect(toastText()).toBe("That is not a pls,fix copy from Excel.");
  });

  it("foreign HTML with real text and no carrier at all falls to the plain-text check", async () => {
    await bootPpt({ local: true });

    pasteBox().dispatchEvent(
      pasteEvent("<p>Quarterly results</p>", "Quarterly results"),
    );
    await settlePpt();

    expect(toastText()).toBe("That is not a pls,fix copy from Excel.");
  });

  it("an empty bundle (valid shape, nothing in it) reads as 'not a copy', never a crash", async () => {
    await bootPpt({ local: true });
    const empty = encodeBundle({ links: [], inbox: [] });

    pasteBox().dispatchEvent(pasteEvent(bundleHtml(empty), ""));
    await settlePpt();

    expect(toastText()).toBe("That is not a pls,fix copy from Excel.");
  });

  it("a bundle from a newer add-in names the reason instead of failing silently", async () => {
    await bootPpt({ local: true });
    const newer = JSON.stringify({
      plsfix: "links",
      v: 2,
      links: [],
      inbox: [],
    });

    pasteBox().dispatchEvent(pasteEvent(bundleHtml(newer), ""));
    await settlePpt();

    expect(toastText()).toBe(
      "This copy comes from a newer pls,fix. Update the add-in.",
    );
  });
});

describe("a paste in relay mode, before local mode has ever been used", () => {
  it("says to switch Settings first, rather than reaching for a store that was never opened", async () => {
    await bootPpt({ local: false });
    const collector = new LocalCollector();
    await exportLinkLocally(collector);

    pasteBox().dispatchEvent(
      pasteEvent(bundleHtml(encodeBundle(collector.bundle())), ""),
    );
    await settlePpt();

    expect(toastText()).toBe("Switch Settings to copy and paste first.");
  });
});

describe("the 'forgets between sessions' hint, and that the pane keeps working without it", () => {
  it("shows the hint by default (this test environment has no IndexedDB) and still ingests a paste", async () => {
    await bootPpt({ local: true });
    const hint = document.getElementById("paste-not-durable")!;
    expect(hint.hidden).toBe(false);

    const collector = new LocalCollector();
    await exportLinkLocally(collector);
    pasteBox().dispatchEvent(
      pasteEvent(bundleHtml(encodeBundle(collector.bundle())), ""),
    );
    await settlePpt();

    expect(toastText()).toContain("waiting in the Inbox");
    expect(
      document.querySelector<HTMLButtonElement>(".inbox-insert"),
    ).toBeTruthy();
  });
});

describe("a paste that partly fails", () => {
  it("carries the failing row's own sentence into Copy details, without stopping the rest", async () => {
    await bootPpt({ local: true });
    const collector = new LocalCollector();
    // Planted first, so each shape's tag predates the export it is about to
    // receive - the order every other insert/update fixture in this suite
    // already uses (assertFresh refuses an update no newer than the tag).
    const good = { id: newId(), token: newToken() };
    plantLink(presentation, good.id, good.token);
    await exportLinkLocally(collector, undefined, good);
    // Sealed with a key that is not this link's own: the shape's real token
    // will never open it, the same shape stress.ppt.links names "a tag
    // edited by hand" / "the wrong key" produces on the relay path.
    const bad = { id: newId(), token: newToken() };
    plantLink(presentation, bad.id, bad.token);
    const wrongKeys = await deriveLinkKeys(newToken());
    await exportLinkLocally(collector, wrongKeys, bad);

    pasteBox().dispatchEvent(
      pasteEvent(bundleHtml(encodeBundle(collector.bundle())), ""),
    );
    await settlePpt();

    expect(toastText()).toBe("Pasted 2 links: 1 updated, 1 failed.");
    // The failing row's own sentence reached the toast's details buffer, not
    // merely a button that happens to render: the label the deck's own
    // failureLine gives a decrypt failure today.
    expect(await copiedDetails()).toBe("Model!B4:F12: open: cannot decrypt");
  });
});
