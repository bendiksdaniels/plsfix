// @vitest-environment jsdom
// The PowerPoint inbox's Slide and Where pickers, driven through the real
// pptpane.html markup: defaults, a chosen slide and spot reaching both Insert
// and "Paste latest linked", and the Slide picker's options following the
// deck's slide count. Split out of inbox.audit.test.ts, already at the
// 400-line cap.

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
import type * as RelayModule from "../link/relay";

enableStrictLoadSemantics();

// Same fake-relay swap inbox.audit.test.ts uses: the pane builds its own
// RelayClient from document.baseURI.
vi.mock("../link/relay", async (importOriginal) => {
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

async function seed(png = fakePng(100, 50)): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload: Payload = {
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
  const keys = await deriveLinkKeys(token);
  await relay.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
  );
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

async function boot(slides = 3): Promise<void> {
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
  const host = installFakePpt({ slides, storage });
  presentation = host.presentation;
  helpers = host.helpers;
  helpers.selectSlide(presentation.slides[0]!.id);
  await import("./main");
  await settle();
}

function click(id: string): void {
  document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
}

function select(id: string): HTMLSelectElement {
  return document.getElementById(id) as HTMLSelectElement;
}

async function settle(rounds = 12): Promise<void> {
  for (let round = 0; round < rounds; round += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function inboxButtons(): HTMLButtonElement[] {
  return [
    ...document.querySelectorAll<HTMLButtonElement>("#inbox-list button"),
  ];
}

afterEach(() => {
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("the insert pickers' defaults", () => {
  beforeEach(async () => {
    await boot();
  });

  it("start on This slide and Free space", () => {
    expect(select("insert-slide").value).toBe("");
    expect(select("insert-where").value).toBe("free");
  });

  it("lists one option per slide, read at boot", () => {
    const options = [...select("insert-slide").options];
    expect(options.map((option) => option.textContent)).toEqual([
      "This slide",
      "Slide 1",
      "Slide 2",
      "Slide 3",
    ]);
  });
});

describe("the pickers drive where an insert lands", () => {
  beforeEach(async () => {
    await boot();
  });

  it("sends an Insert to the chosen slide and spot", async () => {
    await seed();
    click("refresh-inbox");
    await settle();

    select("insert-slide").value = "2";
    select("insert-where").value = "top-right";
    inboxButtons()[0]!.click();
    await settle();

    expect(toastText()).toBe("Inserted Model!B4:F12.");
    const target = presentation.slides[1]!;
    expect(target.shapes).toHaveLength(1);
    // top-right: right of the slide's midline, in the top margin band.
    expect(target.shapes[0]!.left).toBeGreaterThanOrEqual(486);
    expect(target.shapes[0]!.top).toBeGreaterThanOrEqual(36);
    // The view followed an explicitly named slide.
    expect(presentation.selectedSlideIds).toEqual([target.id]);
  });

  it("sends Paste latest linked to the same chosen target", async () => {
    await seed();

    select("insert-slide").value = "3";
    select("insert-where").value = "whole";
    click("paste-latest-linked");
    await settle();

    expect(toastText()).toContain("Inserted");
    expect(presentation.slides[2]!.shapes).toHaveLength(1);
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("keeps This slide and Free space acting exactly as before the pickers", async () => {
    await seed();
    click("refresh-inbox");
    await settle();

    inboxButtons()[0]!.click();
    await settle();

    expect(toastText()).toBe("Inserted Model!B4:F12.");
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
  });
});

describe("the Slide picker follows the deck's slide count", () => {
  beforeEach(async () => {
    await boot();
  });

  it("grows when a refresh finds a new slide", async () => {
    presentation.addSlide();
    click("refresh-inbox");
    await settle();

    const options = [...select("insert-slide").options];
    expect(options.map((option) => option.textContent)).toEqual([
      "This slide",
      "Slide 1",
      "Slide 2",
      "Slide 3",
      "Slide 4",
    ]);
  });

  it("falls back to This slide alone when the read fails", async () => {
    helpers.failNextSync(new Error("PowerPoint stopped answering."));
    click("refresh-inbox");
    await settle();

    expect([...select("insert-slide").options]).toHaveLength(1);
  });
});
