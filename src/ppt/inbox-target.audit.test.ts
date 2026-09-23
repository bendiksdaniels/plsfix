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
import { settlePpt, trackPptBoot } from "../../test/ppt-ready";
import type * as RelayModule from "../link/relay";
import { SLIDE_MARGIN } from "./placement";

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
  const booted = trackPptBoot();
  await import("./main");
  await booted;
  await settle();
}

function click(id: string): void {
  document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
}

function select(id: string): HTMLSelectElement {
  return document.getElementById(id) as HTMLSelectElement;
}

async function settle(rounds = 12): Promise<void> {
  await settlePpt(rounds);
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function inboxButtons(): HTMLButtonElement[] {
  return [
    ...document.querySelectorAll<HTMLButtonElement>("#inbox-list button"),
  ];
}

function isError(): boolean {
  return document.getElementById("toast")?.className.includes("error") ?? true;
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
    expect(target.shapes[0]!.top).toBeGreaterThanOrEqual(SLIDE_MARGIN);
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

  it("sends Insert to a selected placeholder and consumes it", async () => {
    const item = await seed();
    click("refresh-inbox");
    await settle();
    const placeholder = presentation.addShape(presentation.slides[0]!, {
      type: "Placeholder",
      hasText: false,
      left: 100,
      top: 80,
      width: 300,
      height: 150,
    });
    helpers.selectShapes([placeholder.id]);
    select("insert-where").value = "selected-shape";

    inboxButtons()[0]!.click();
    await settle();

    expect(toastText()).toBe(`Inserted ${item.label}.`);
    const ids = presentation.slides[0]!.shapes.map((one) => one.id);
    expect(ids).not.toContain(placeholder.id);
    expect(ids).toHaveLength(1);
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

  it("falls back to This slide alone when the read fails, and says so quietly", async () => {
    helpers.failNextSync(new Error("PowerPoint stopped answering."));
    click("refresh-inbox");
    await settle();

    expect([...select("insert-slide").options]).toHaveLength(1);
    // The inbox itself still refreshed fine, so the toast stays a success -
    // but the slide-count failure is not swallowed: it lands in the details
    // a "Copy details" button exposes, the same quiet-failure path
    // refreshQuietly already uses for the Links list.
    expect(isError()).toBe(false);
    expect(document.querySelector(".toast-copy")).not.toBeNull();
  });
});

describe("a Slide pick that has gone stale", () => {
  beforeEach(async () => {
    await boot();
  });

  // The picker only ever learns a position (target.ts), so a deck that lost
  // a slide after the picker was last rendered must answer the pane's own
  // sentence at the moment of the click, never a raw office.js error.
  it("answers plainly, not a raw error, when the picked slide is gone by click time", async () => {
    await seed();
    click("refresh-inbox");
    await settle();
    select("insert-slide").value = "3";
    presentation.slides.pop();

    inboxButtons()[0]!.click();
    await settle();

    expect(toastText()).toBe("Slide 3 is gone: pick a slide again.");
    expect(isError()).toBe(true);
    // The item is still waiting, and every button is live again.
    expect(inboxButtons()).toHaveLength(1);
    expect(inboxButtons()[0]?.disabled).toBe(false);
  });
});
