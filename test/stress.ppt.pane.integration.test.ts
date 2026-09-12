// @vitest-environment jsdom
// Stress pass on the PowerPoint pane itself: every button pressed twice, at
// the wrong time and while a batch is still running, against a relay that
// answers badly. Drives the shipped pptpane.html through src/ppt/main.ts.
// Invariant: one action runs at a time, every press ends in a sentence, and
// the pane always comes back usable.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
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
} from "../src/link/model";
import {
  createWorkspace,
  WORKSPACE_STORAGE_KEY,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import type * as RelayModule from "../src/link/relay";
import { RelayError } from "../src/link/relay";

enableStrictLoadSemantics();

// The pane builds its own RelayClient from document.baseURI, so the fake takes
// that constructor's place; every other export of the module stays real.
vi.mock("../src/link/relay", async (importOriginal) => {
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

// A link already on slide 2, the way a colleague's deck arrives.
async function plant(): Promise<string> {
  const id = newId();
  const token = newToken();
  await publish(id, token, picture(fakePng(800, 400)));
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
  return shape.id;
}

// One export waiting in the Inbox for this workspace.
async function waiting(): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  await publish(id, token, picture(fakePng(200, 100)));
  const item: InboxItem = {
    id,
    token,
    kind: "range",
    label: "Model!B4:F12",
    src: SRC,
    createdAt: new Date().toISOString(),
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
  const storage = new Map([[WORKSPACE_STORAGE_KEY, workspace.exportKey]]);
  const host = installFakePpt({ slides: 3, storage });
  presentation = host.presentation;
  helpers = host.helpers;
  helpers.selectSlide(presentation.slides[0]!.id);
  await import("../src/ppt/main");
  await settle();
}

function click(id: string): void {
  button(id).click();
}
function button(id: string): HTMLButtonElement {
  return document.getElementById(id) as HTMLButtonElement;
}
async function settle(rounds = 12): Promise<void> {
  for (let round = 0; round < rounds; round += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}
function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}
function hasDetails(): boolean {
  return document.querySelector("#toast .toast-copy") !== null;
}
function linkRows(): HTMLTableRowElement[] {
  return [...document.querySelectorAll<HTMLTableRowElement>("#link-rows tr")];
}
function tickRow(index: number): void {
  const box = linkRows()[index]?.querySelector("input");
  if (!box) throw new Error(`no row ${String(index)} to tick`);
  box.checked = true;
  box.dispatchEvent(new Event("change"));
}
function pressEnterOnKeyField(): void {
  const field = document.getElementById("workspace-key") as HTMLInputElement;
  field.value = workspace.exportKey;
  field.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
  );
}

// A relay call the test releases by hand, so a batch can be caught mid-flight.
function heldFetch(): () => void {
  let release = (): void => undefined;
  const gate = new Promise<void>((done) => {
    release = () => done();
  });
  const real = relay.fetchLinks.bind(relay);
  vi.spyOn(relay, "fetchLinks").mockImplementation(async (items) => {
    await gate;
    return real(items);
  });
  return release;
}

afterEach(() => {
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("a second action while the first is still running", () => {
  it("refuses the link key's Enter mid-batch and keeps the pane busy", async () => {
    await boot();
    await plant();
    click("refresh-links");
    await settle();
    const release = heldFetch();

    click("update-all");
    await settle(3);
    expect(button("update-all").disabled).toBe(true);

    // The field is not a button, so setBusy never disabled it: pressing Enter
    // here used to start a second flow whose own finish re-enabled every
    // button while the batch was still in the host.
    pressEnterOnKeyField();
    await settle();

    expect(toastText()).toBe("Wait for the last action to finish.");
    expect(button("update-all").disabled).toBe(true);
    expect(button("break-selected").disabled).toBe(true);

    release();
    await settle();
    expect(toastText()).toBe("1 up to date");
    expect(button("update-all").disabled).toBe(false);
  });

  it("disables every button while a batch runs and gives them all back", async () => {
    await boot();
    await plant();
    click("refresh-links");
    await settle();
    const release = heldFetch();

    click("update-all");
    await settle(3);
    const ids = [
      "refresh-links",
      "update-selected",
      "update-slide",
      "revert-selected",
      "break-selected",
      "go-to-slide",
      "refresh-inbox",
      "paste-latest-linked",
      "save-key",
      "forget-key",
      "align-objects",
    ];
    expect(ids.filter((id) => !button(id).disabled)).toEqual([]);

    release();
    await settle();
    expect(ids.filter((id) => button(id).disabled)).toEqual([]);
  });

  it("inserts once when Paste latest linked is pressed twice", async () => {
    await boot();
    await waiting();
    click("refresh-inbox");
    await settle();

    click("paste-latest-linked");
    click("paste-latest-linked");
    await settle();

    expect(toastText()).toBe("Inserted Model!B4:F12.");
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(linkRows()).toHaveLength(1);
  });

  it("breaks once when Break is pressed twice", async () => {
    await boot();
    await plant();
    click("refresh-links");
    await settle();
    tickRow(0);

    click("break-selected");
    click("break-selected");
    await settle();

    expect(toastText()).toBe("1 link broken. The picture stays on the slide.");
    expect(linkRows()).toHaveLength(0);
    expect(presentation.slides[1]!.shapes).toHaveLength(1);
  });
});

describe("a button pressed at the wrong moment", () => {
  it("asks for a slide when nothing is selected in the deck", async () => {
    await boot();
    await waiting();
    click("refresh-inbox");
    await settle();
    helpers.clearSelection();

    click("paste-latest-linked");
    await settle();

    expect(toastText()).toBe("insert Model!B4:F12: select a slide first.");
    expect(button("paste-latest-linked").disabled).toBe(false);
    // The export is still waiting: nothing was consumed by the refusal.
    expect(await relay.listInbox(workspace.id, workspace.auth)).toHaveLength(1);
  });

  it("asks for a tick before Update selected, Revert, Break and Go to slide", async () => {
    await boot();
    await plant();
    click("refresh-links");
    await settle();

    for (const [id, sentence] of [
      ["update-selected", "Tick a link in the list first."],
      ["revert-selected", "Tick the rows to revert."],
      ["break-selected", "Tick a link in the list first."],
      ["go-to-slide", "Tick a link in the list first."],
    ] as [string, string][]) {
      click(id);
      await settle();
      expect(toastText()).toBe(sentence);
      expect(button(id).disabled).toBe(false);
    }
  });

  it("says which slide is missing links when the wrong one is selected", async () => {
    await boot();
    await plant();
    click("refresh-links");
    await settle();
    helpers.selectSlide(presentation.slides[2]!.id);

    click("update-slide");
    await settle();

    expect(toastText()).toBe("No links on this slide");
    expect(button("update-slide").disabled).toBe(false);
  });

  it("goes to a slide the deck no longer holds without a raw error", async () => {
    await boot();
    await plant();
    click("refresh-links");
    await settle();
    tickRow(0);
    presentation.slides.splice(1, 1);

    click("go-to-slide");
    await settle();

    expect(toastText()).not.toMatch(/ItemNotFound|Exception|undefined/);
    expect(button("go-to-slide").disabled).toBe(false);
  });
});

describe("a relay that answers badly", () => {
  it("counts a rate-limited update as failed and keeps the reason to copy", async () => {
    await boot();
    await plant();
    click("refresh-links");
    await settle();
    for (const call of ["fetchLinks", "getLink"] as const) {
      vi.spyOn(relay, call).mockRejectedValue(
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

  it("says an export is too big to send in words", async () => {
    await boot();
    await waiting();
    click("refresh-inbox");
    await settle();
    vi.spyOn(relay, "getLink").mockRejectedValue(
      new RelayError("tooLarge", "relay GET /api/links: 413", 413),
    );

    click("paste-latest-linked");
    await settle();

    expect(toastText()).toContain("too big to send");
    expect(button("paste-latest-linked").disabled).toBe(false);
  });

  // REPORTED (slice K3 owns src/ppt/links.ts): a 429, 507 or 500 travels to
  // the modeller as the relay's own HTTP sentence, while a 413 is reworded by
  // TOO_LARGE. Fix: give `failureLine` and `insertFromInbox` in
  // src/ppt/links.ts a wording per RelayError kind/status - "the relay is
  // busy, try again in a minute" for 429 and "the relay is full" for 507 -
  // the way TOO_LARGE already answers a 413.
  it.skip("says a rate-limited relay is busy rather than showing 429", async () => {
    await boot();
    await waiting();
    click("refresh-inbox");
    await settle();
    vi.spyOn(relay, "getLink").mockRejectedValue(
      new RelayError("server", "relay GET /api/links: 429", 429),
    );

    click("paste-latest-linked");
    await settle();

    expect(toastText()).not.toMatch(/429|\/api\//);
  });

  // REPORTED (slice K3 owns src/ppt/links.ts): insertFromInbox deletes the
  // inbox row after the shape is on the slide, and a delete that fails with
  // anything but a 404 throws - so the insert reports failure with the shape
  // already inserted, the item stays in the list, and a second press lands a
  // duplicate. Fix: in src/ppt/links.ts, treat a failed deleteInbox as a note
  // on a successful insert (the row expires on its own after 7 days).
  it.skip("keeps a landed insert a success when the inbox row cannot be dropped", async () => {
    await boot();
    await waiting();
    click("refresh-inbox");
    await settle();
    vi.spyOn(relay, "deleteInbox").mockRejectedValue(
      new RelayError("server", "relay DELETE /api/inbox: 500", 500),
    );

    click("paste-latest-linked");
    await settle();

    expect(toastText()).toContain("Inserted");
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
  });
});
