// @vitest-environment jsdom
// Hunt pass 1: reloading the pane mid-action - the JS context is discarded
// and the in-flight PowerPoint.run() abandoned, never rejected. A generic
// hung sync never corrupts a fresh boot; a Change source interrupted between
// its retag and its repaint reads update-available on the reloaded list.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../../src/link/crypto";
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
import type * as RelayModule from "../../src/link/relay";
import { TRANSPORT_STORAGE_KEY } from "../../src/link/transport-setting";
import {
  createWorkspace,
  WORKSPACE_STORAGE_KEY,
} from "../../src/link/workspace";
import { FakeRelay } from "../fakerelay";
import { fakePng } from "../fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePresentation,
  type FakePptHelpers,
} from "../fakeppt";
import { settlePpt, trackPptBoot } from "../ppt-ready";

enableStrictLoadSemantics();

const SRC = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

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

function newId(): string {
  return newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
}

interface Carry {
  presentation?: FakePresentation;
  storage?: Map<string, string>;
  relay?: FakeRelay;
}

// The three things a real reload never touches: the deck itself, this
// device's OfficeRuntime.storage (the pairing key lives there, not in JS
// heap) and the server this pane talks to. Only the module graph resets -
// carrying none of the three over is exactly what made the old version of
// this test unable to tell a correct reload from a broken one.
async function bootPpt(
  carry: Carry = {},
): Promise<{ helpers: FakePptHelpers; presentation: FakePresentation }> {
  vi.resetModules();
  uninstallFakePpt();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  relay = carry.relay ?? new FakeRelay();
  const installed = installFakePpt({
    slides: 3,
    presentation: carry.presentation,
    storage: carry.storage,
  });
  const booted = trackPptBoot();
  await import("../../src/ppt/main");
  await booted;
  await settlePpt();
  return installed;
}

afterEach(() => {
  uninstallFakePpt();
});

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function click(id: string): void {
  document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
}

function statusText(): string {
  return document.querySelector("#link-rows .link-status")?.textContent ?? "";
}

// PowerPoint.run reads the CURRENT global at call time, not at the time the
// click handler started, and a click's own promise chain crosses several
// real awaits (WebCrypto, the relay) before it ever reaches one - far more
// than the synchronous prefix of a second bootPpt() (which swaps the global
// straight away). Left alone, the "reload" below would race ahead and the
// abandoned action would run its retag AND its repaint against the FRESH
// runtime instead of the old one, hanging nothing. Waiting for the old
// runtime's own sync count to reach the hung sync's ordinal is what a real
// reload cannot do (there is no old runtime left to ask) but is the only way
// this fake can prove the retag truly landed before the reload, not after.
async function untilSyncCount(
  helpers: FakePptHelpers,
  target: number,
): Promise<void> {
  for (let turn = 0; turn < 500 && helpers.syncCount() < target; turn += 1) {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  }
}

describe("a reload while an action is still in flight", () => {
  it("abandoning the old instance's hung sync does not stop a fresh reload from working", async () => {
    const capturedErrors: string[] = [];
    const onError = (event: ErrorEvent): void => {
      capturedErrors.push(`error: ${String(event.error ?? event.message)}`);
    };
    const onRejection = (event: PromiseRejectionEvent): void => {
      capturedErrors.push(`unhandledrejection: ${String(event.reason)}`);
    };
    window.addEventListener("error", onError);
    window.addEventListener("unhandledrejection", onRejection);

    const firstHelpers = (await bootPpt()).helpers;
    firstHelpers.hangNextSync();
    // Fire-and-forget, exactly like a real click: main.ts's act() returns
    // void. This promise chain is never awaited, settled or referenced
    // again - the old JS context abandoning it is the whole scenario.
    click("refresh-links");

    // The "reload": a second pane, a fresh module graph (vi.resetModules
    // inside bootPpt), a fresh deck. Nothing from above carries over.
    await bootPpt();

    expect(toastText()).toBe("");
    expect(
      document.querySelector<HTMLButtonElement>("#refresh-links")?.disabled,
    ).toBe(false);

    click("refresh-links");
    await settlePpt();

    expect(toastText()).toBe("No linked objects in this deck.");
    expect(capturedErrors).toEqual([]);

    window.removeEventListener("error", onError);
    window.removeEventListener("unhandledrejection", onRejection);
  });
});

describe("a reload abandons a Change source between its retag and its repaint", () => {
  it("the reloaded list reads update-available, never falsely current, and Update all repaints it", async () => {
    // Paired and in relay mode: both live in OfficeRuntime.storage, which
    // (unlike a JS variable) survives a real reload, so the second boot
    // must be handed the SAME map to come up paired against the SAME deck.
    const ws = await createWorkspace({
      get: async () => null,
      set: async () => undefined,
      remove: async () => undefined,
    });
    const storage = new Map<string, string>();
    storage.set(WORKSPACE_STORAGE_KEY, ws.exportKey);
    storage.set(TRANSPORT_STORAGE_KEY, "relay");

    const first = await bootPpt({ storage });

    // The link already on the deck, tracking Model_v4 - planted directly
    // (the shortcut every other hunt file uses for an existing shape)
    // rather than through Insert, which this scenario has no need of.
    const oldToken = newToken();
    const oldKeys = await deriveLinkKeys(oldToken);
    const oldId = newId();
    const oldPng = fakePng(200, 100);
    const oldPayload: Payload = {
      v: 1,
      kind: "picture",
      mime: "image/png",
      width: 800,
      height: 400,
      png: oldPng,
      src: SRC,
      pushedAt: new Date().toISOString(),
      hash: "0".repeat(64),
    };
    await relay.putLink(
      oldId,
      oldKeys.auth,
      await seal(oldKeys.enc, oldId, encodePayload(oldPayload)),
    );
    const shape = first.presentation.addShape(first.presentation.slides[0]!, {
      type: "GeometricShape",
      fillImage: oldPng,
      left: 100,
      top: 80,
      width: 400,
      height: 200,
    });
    const oldTag: LinkTag = {
      v: 1,
      id: oldId,
      kind: "range",
      rev: 1,
      src: SRC,
      pushedAt: oldPayload.pushedAt,
    };
    shape.tags.set(TAG_LINK, encodeTag(oldTag));
    shape.tags.set(TAG_KEY, oldToken);

    // The newer export waiting in the Inbox: Change source's one candidate.
    const newerToken = newToken();
    const newKeys = await deriveLinkKeys(newerToken);
    const newerId = newId();
    const newPng = fakePng(400, 200);
    const newPayload: Payload = {
      v: 1,
      kind: "picture",
      mime: "image/png",
      width: 1600,
      height: 800,
      png: newPng,
      src: { ...SRC, workbook: "Model_v5.xlsx" },
      pushedAt: new Date().toISOString(),
      hash: "1".repeat(64),
    };
    await relay.putLink(
      newerId,
      newKeys.auth,
      await seal(newKeys.enc, newerId, encodePayload(newPayload)),
    );
    const inboxItem: InboxItem = {
      id: newerId,
      token: newerToken,
      kind: "range",
      label: "Model!B4:F12",
      src: newPayload.src,
      createdAt: newPayload.pushedAt,
    };
    await relay.postInbox(
      ws.id,
      ws.auth,
      newerId,
      await seal(ws.enc, ws.id, encodeInboxItem(inboxItem)),
    );

    // Rescan the deck and the Inbox so both are on screen for the picker.
    click("refresh-links");
    await settlePpt();
    click("refresh-inbox");
    await settlePpt();

    const checkbox = document.querySelector<HTMLInputElement>(
      "#link-rows input[type=checkbox]",
    );
    checkbox!.checked = true;
    checkbox!.dispatchEvent(new Event("change", { bubbles: true }));
    click("change-source");
    await settlePpt();
    const list = document.getElementById(
      "change-source-list",
    ) as HTMLSelectElement;
    list.value = newerId;

    // Hang the retag's OWN sync (the very next one). The fake, like real
    // Office.js, applies a queued write (tags.add, fill.setImage) the moment
    // it is called, not when the sync that reports it settles - so hanging
    // the retag's sync still lands its write and, because changeSource never
    // gets past an `await` that never resolves, the repaint that would
    // follow never starts at all: exactly "the pane gone after the retag
    // lands and before the repaint". Fire-and-forget, never awaited - a
    // reload right here is what abandons it mid-flight.
    const base = first.helpers.syncCount();
    first.helpers.hangNextSync();
    click("change-source-confirm");
    // Wait for the retag to actually reach the OLD runtime (see
    // untilSyncCount) before the reload swaps it out from under it -
    // otherwise the action runs, unhung, against the NEW one instead.
    await untilSyncCount(first.helpers, base + 1);

    // The "reload": a second pane on the SAME deck, storage and relay - a
    // fresh JS context is the only thing that actually resets.
    await bootPpt({ presentation: first.presentation, storage, relay });

    // Fix: retag with rev 0, not the real one, so a shape frozen between
    // retag and repaint still reads as stale. Before the fix the tag already
    // carried the relay's own rev, so this read "Up to date" while the
    // picture was still Model_v4's.
    expect(statusText()).toBe("Update available");
    expect(shape.fillImage).toBe(oldPng);

    click("update-all");
    await settlePpt();

    expect(shape.fillImage).toBe(newPng);
    expect(statusText()).toBe("Up to date");
  });
});
