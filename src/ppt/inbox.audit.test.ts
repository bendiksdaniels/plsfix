// @vitest-environment jsdom
// The PowerPoint pane's Inbox and Settings, driven through the real
// pptpane.html markup: one-click paste, the Insert button on a row, the
// overlap note, the pairing key going in and coming out, and the busy flag
// over a list the pane redraws while an action is still running.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../link/crypto";
import {
  encodeInboxItem,
  encodePayload,
  newLinkId,
  type InboxItem,
  type LinkKind,
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
import { RelayError } from "../link/relay";

enableStrictLoadSemantics();

// The pane builds its own RelayClient from document.baseURI, so the fake takes
// that constructor's place; every other export of the module stays real.
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

function picture(png: string, workbook = SRC.workbook): Payload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png,
    src: { ...SRC, workbook },
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
async function seed(
  png = fakePng(800, 400),
  workbook = SRC.workbook,
  kind: LinkKind = "range",
): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload = picture(png, workbook);
  await publish(id, token, payload);
  const item: InboxItem = {
    id,
    token,
    kind,
    label: kind === "chart" ? "Charts: Revenue" : `${SRC.sheet}!${SRC.ref}`,
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

async function boot(paired = true): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  relay = new FakeRelay();
  workspace = await createWorkspace(noStore);
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  const storage = new Map<string, string>();
  if (paired) storage.set(WORKSPACE_STORAGE_KEY, workspace.exportKey);
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

async function settle(rounds = 12): Promise<void> {
  for (let round = 0; round < rounds; round += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
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

afterEach(() => {
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("the Inbox and the pairing key", () => {
  beforeEach(async () => {
    await boot();
  });

  it("pastes the newest waiting export in one click, and says so when none waits", async () => {
    click("paste-latest-linked");
    await settle();
    expect(toastText()).toBe(
      "Nothing waiting from Excel. Export an object there first.",
    );

    await seed(fakePng(100, 50));
    relay.now += 1;
    const newest = await seed(fakePng(200, 100), "Model_v9.xlsx");
    click("paste-latest-linked");
    await settle();

    expect(toastText()).toBe(`Inserted ${newest.label}.`);
    expect(presentation.slides[0]!.shapes[0]!.fillImage).toBe(
      fakePng(200, 100),
    );
    // The one it pasted leaves the list; the older export still waits.
    expect(inboxButtons()).toHaveLength(1);
  });

  it("says the slide was full when the insert had to cover something", async () => {
    presentation.addShape(presentation.slides[0]!, {
      left: 0,
      top: 0,
      width: 960,
      height: 540,
    });
    await seed();
    click("paste-latest-linked");
    await settle();

    expect(toastText()).toBe(
      "Inserted Model!B4:F12. Placed over other objects: no free space on this slide",
    );
  });

  it("refuses an insert with no slide selected and stays usable", async () => {
    await seed();
    click("refresh-inbox");
    await settle();
    helpers.clearSelection();

    inboxButtons()[0]?.click();
    await settle();

    expect(toastText()).toBe("insert Model!B4:F12: select a slide first.");
    expect(isError()).toBe(true);
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
    // The item is still waiting, and every button is live again.
    expect(inboxButtons()).toHaveLength(1);
    expect(inboxButtons()[0]?.disabled).toBe(false);
  });

  it("forgets a key and takes a new one, from the field and from Enter", async () => {
    const field = document.getElementById("workspace-key") as HTMLInputElement;
    const state = document.getElementById("workspace-state");
    expect(state?.textContent).toBe("Paired");

    click("forget-key");
    await settle();
    expect(toastText()).toBe(
      "Link key forgotten. The links already in this deck still update.",
    );
    expect(state?.textContent).toContain("Not paired");

    click("save-key");
    await settle();
    expect(toastText()).toBe("Paste the link key from Excel first.");

    // Anything that is not the key itself, a pasted "Link key: ..." label
    // included, gets the same sentence rather than a codec error.
    for (const typed of [
      "not a link key",
      `Link key: ${workspace.exportKey}`,
    ]) {
      field.value = typed;
      click("save-key");
      await settle();
      expect(toastText()).toBe(
        "That is not a link key. Copy it again from Excel.",
      );
    }

    // A key pasted out of an e-mail, with the whitespace it was wrapped in.
    field.value = `\n  ${workspace.exportKey}\t\n`;
    field.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    await settle();
    expect(toastText()).toBe("Paired with Excel.");
    expect(state?.textContent).toBe("Paired");
    // The key is the secret: it is stored, never echoed back into the field.
    expect(field.value).toBe("");
  });

  // Pairing with another workspace must not leave the previous one's exports
  // on screen: their Insert buttons still work, and they belong to a key this
  // deck no longer holds.
  it("drops the old workspace's exports when a new key is pasted", async () => {
    await seed();
    click("refresh-inbox");
    await settle();
    expect(inboxButtons()).toHaveLength(1);

    const other = await createWorkspace({
      get: async () => null,
      set: async () => undefined,
      remove: async () => undefined,
    });
    vi.spyOn(relay, "listInbox").mockRejectedValue(
      new RelayError("network", "relay GET /api/inbox: network error"),
    );
    const field = document.getElementById("workspace-key") as HTMLInputElement;
    field.value = other.exportKey;
    click("save-key");
    await settle();

    expect(toastText()).toBe("Paired with Excel.");
    expect(inboxButtons()).toHaveLength(0);
  });

  // The Inbox list is rebuilt in the middle of an insert - the item it just
  // consumed has to leave it - and the buttons that rebuild carries are new
  // ones the busy flag never saw. A second Insert must not be clickable while
  // the first is still working.
  it("keeps the redrawn Insert buttons disabled while an insert runs", async () => {
    await seed();
    await seed(fakePng(100, 50), "Other.xlsx");
    click("refresh-inbox");
    await settle();
    expect(inboxButtons()).toHaveLength(2);

    // The rescan that follows the insert never answers, so the pane is still
    // busy when the assertions run.
    vi.spyOn(relay, "status").mockReturnValue(new Promise(() => undefined));
    inboxButtons()[0]!.click();
    await settle();

    expect(button("refresh-inbox").disabled).toBe(true);
    expect(inboxButtons().map((one) => one.disabled)).toEqual([true]);
  });

  // Only the refusal the user can act on is reworded; everything else reaches
  // the toast exactly as the relay reported it.
  it("keeps the relay's own sentence when an insert cannot reach it", async () => {
    await seed();
    click("refresh-inbox");
    await settle();
    vi.spyOn(relay, "getLink").mockRejectedValue(
      new RelayError("network", "relay GET /api/links/aaaa: network error"),
    );

    inboxButtons()[0]!.click();
    await settle();

    expect(toastText()).toBe("relay GET /api/links/aaaa: network error");
    expect(isError()).toBe(true);
    expect(inboxButtons()).toHaveLength(1);
    expect(button("refresh-inbox").disabled).toBe(false);
  });

  // The relay stamps an inbox row in whole seconds, so two exports pushed in
  // the same second arrive tied. The pane's sort must not shuffle them: the
  // order the relay answered in is the tie-break, and it is stable.
  it("keeps the relay's order between exports of the same second", async () => {
    const first = await seed(fakePng(100, 50), "A.xlsx");
    const second = await seed(fakePng(200, 100), "B.xlsx");
    expect(relay.now).toBe(1_000_000);

    click("refresh-inbox");
    await settle();

    expect(inboxButtons().map((one) => one.getAttribute("aria-label"))).toEqual(
      [`Insert ${first.label}`, `Insert ${second.label}`],
    );
    click("paste-latest-linked");
    await settle();
    expect(toastText()).toBe(`Inserted ${first.label}.`);
  });

  // The only route below PowerPointApi 1.8: the picture goes in through the
  // Office selection API. A host that refuses it must leave the pane usable
  // and the export still waiting.
  it("inserts through the selection below 1.8, and clears busy when it fails", async () => {
    helpers.setSupported(
      (set, version) => set === "PowerPointApi" && Number(version) <= 1.5,
    );
    await seed();
    click("refresh-inbox");
    await settle();

    helpers.failNextSelectionInsert("the host is out of memory");
    inboxButtons()[0]!.click();
    await settle();

    expect(toastText()).toBe("insert Model!B4:F12: the host is out of memory");
    expect(isError()).toBe(true);
    expect(button("refresh-inbox").disabled).toBe(false);
    expect(inboxButtons()).toHaveLength(1);
    expect(inboxButtons()[0]!.disabled).toBe(false);

    inboxButtons()[0]!.click();
    await settle();
    expect(toastText()).toBe("Inserted Model!B4:F12.");
    expect(presentation.slides[0]!.shapes[0]!.type).toBe("Image");
  });
});
