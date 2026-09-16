// @vitest-environment jsdom
// J audit: withSyncDeadline generalised past chart-draw.ts to every other
// PowerPoint.run sync in the pane, so a plain picture insert whose own round
// trip the host swallows must free the busy pane and toast the sentence
// instead of leaving every button disabled for ever (lessons, 30.08 and
// 08.09). Driven through the real pptpane.html markup and main.ts, the same
// rig as inbox.audit.test.ts.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../link/crypto";
import {
  encodeInboxItem,
  encodePayload,
  newLinkId,
  type InboxItem,
  type Payload,
} from "../link/model";
import {
  createWorkspace,
  WORKSPACE_STORAGE_KEY,
  type KeyStore,
  type Workspace,
} from "../link/workspace";
import { FakeRelay } from "../../test/fakerelay";
import { fakePng } from "../../test/fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "../../test/fakeppt";
import { settleUntil } from "../../test/hung-sync";
import type * as RelayModule from "../link/relay";

enableStrictLoadSemantics();

const SRC = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

let relay: FakeRelay;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let workspace: Workspace;

const noStore: KeyStore = {
  get: async () => null,
  set: async () => undefined,
  remove: async () => undefined,
};

function newId(): string {
  return newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
}

function picture(png: string): Payload {
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

async function publish(
  id: string,
  token: string,
  payload: Payload,
): Promise<void> {
  const keys = await deriveLinkKeys(token);
  await relay.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
  );
}

// One export waiting in the pane's inbox, sealed the way Excel seals it.
async function seed(): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload = picture(fakePng(800, 400));
  await publish(id, token, payload);
  const item: InboxItem = {
    id,
    token,
    kind: "range",
    label: `${SRC.sheet}!${SRC.ref}`,
    src: payload.src,
    createdAt: payload.pushedAt,
  };
  await relay.postInbox(
    workspace.id,
    workspace.auth,
    id,
    await seal(workspace.enc, workspace.id, encodeInboxItem(item)),
  );
  return item;
}

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  relay = new FakeRelay();
  workspace = await createWorkspace(noStore);
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  const storage = new Map<string, string>();
  storage.set(WORKSPACE_STORAGE_KEY, workspace.exportKey);
  const host = installFakePpt({ slides: 3, storage });
  presentation = host.presentation;
  helpers = host.helpers;
  helpers.selectSlide(presentation.slides[0]!.id);
  await import("./main");
  await settle();
}

function click(id: string): void {
  document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
}

function button(id: string): HTMLButtonElement {
  return document.getElementById(id) as HTMLButtonElement;
}

// Drains whatever microtasks are pending: real setTimeout(0) rounds while the
// clock is real (boot, seeding), vi.advanceTimersByTimeAsync(0) rounds once
// the test has switched to fake timers to drive the sync deadline.
async function settle(rounds = 12): Promise<void> {
  for (let round = 0; round < rounds; round += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

async function settleFake(rounds = 12): Promise<void> {
  for (let round = 0; round < rounds; round += 1) {
    await vi.advanceTimersByTimeAsync(0);
  }
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function isError(): boolean {
  return document.getElementById("toast")?.className.includes("error") ?? true;
}

function inboxButtons(): HTMLButtonElement[] {
  return [
    ...document.querySelectorAll<HTMLButtonElement>("#inbox-list button"),
  ];
}

// vi.mock is hoisted, so a mocked RelayClient is wired before main.ts's own
// construction of it runs on import - the same trick inbox.audit.test.ts
// uses to hand the pane a fake relay without changing its constructor call.
vi.mock("../link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return relay;
    },
  };
});

afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("an insert the host never answers", () => {
  beforeEach(async () => {
    await boot();
  });

  it("frees the busy pane and toasts the sentence, then inserts normally on retry", async () => {
    await seed();
    click("refresh-inbox");
    await settle();
    expect(inboxButtons()).toHaveLength(1);

    // An empty slide's insert is three round trips: which slide is selected,
    // the boxes already on it, then the insert's own sync - the one that
    // never answers here.
    helpers.hangNextSync(2);
    vi.useFakeTimers();
    inboxButtons()[0]!.click();
    await settleFake();
    // The guard's setBusy(true) ran before the hung sync, same as any other
    // action: every button, including the one just clicked, is disabled.
    expect(button("refresh-inbox").disabled).toBe(true);
    expect(inboxButtons()[0]!.disabled).toBe(true);

    // The deadline is armed only once the item is open, which on a slow machine
    // takes real time: drive the clock until the sentence lands (test/hung-sync.ts).
    await settleUntil(() => toastText().startsWith("PowerPoint stopped"));
    await settleFake();

    expect(toastText()).toBe(
      "PowerPoint stopped answering while inserting the picture",
    );
    expect(isError()).toBe(true);
    // The rejection reached the guard's finally, which clears busy: the pane
    // is not stuck the way a swallowed batch used to leave it.
    expect(button("refresh-inbox").disabled).toBe(false);
    expect(inboxButtons()).toHaveLength(1);
    expect(inboxButtons()[0]!.disabled).toBe(false);
    // Nothing was left on the slide: the picture's id was never confirmed.
    expect(presentation.slides[0]!.shapes).toHaveLength(0);

    // The pane is not jammed: the same export inserts normally right after,
    // since the failed attempt never reached the deleteInbox that follows a
    // successful one.
    vi.useRealTimers();
    inboxButtons()[0]!.click();
    await settle();
    expect(toastText()).toBe("Inserted Model!B4:F12.");
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
  });
});
