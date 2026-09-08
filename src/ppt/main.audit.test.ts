// @vitest-environment jsdom
// The PowerPoint pane as the user drives it: every button in pptpane.html
// clicked against the fake deck and a fake relay, so the wiring, the toast
// sentence and the busy flag are covered rather than only the flows below
// them. The relay client is replaced by the in-memory FakeRelay every other
// ppt.* suite runs against; nothing else is mocked.

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
): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload = picture(png, workbook);
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

// The pane boots on import, so the markup, the deck and the stored pairing key
// all go in first. `paired` is what a deck that has never seen a key models.
async function boot(paired = true): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  relay = new FakeRelay();
  workspace = await createWorkspace({
    get: async () => null,
    set: async () => undefined,
    remove: async () => undefined,
  });
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

function linkRows(): HTMLTableRowElement[] {
  return [...document.querySelectorAll<HTMLTableRowElement>("#link-rows tr")];
}

function inboxButtons(): HTMLButtonElement[] {
  return [
    ...document.querySelectorAll<HTMLButtonElement>("#inbox-list button"),
  ];
}

function tickRow(index: number): void {
  const box = linkRows()[index]?.querySelector("input");
  if (!box) throw new Error(`no row ${String(index)} to tick`);
  box.checked = true;
  box.dispatchEvent(new Event("change"));
}

// One export inserted on the active slide, through the pane's own Insert.
async function insertOne(item?: InboxItem): Promise<InboxItem> {
  const waiting = item ?? (await seed());
  click("refresh-inbox");
  await settle();
  inboxButtons().at(-1)?.click();
  await settle();
  return waiting;
}

afterEach(() => {
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("the Links tab", () => {
  beforeEach(async () => {
    await boot();
  });

  it("lists the deck's links, ticks one and reports the update", async () => {
    const item = await insertOne();
    expect(toastText()).toBe("Inserted Model!B4:F12.");
    expect(linkRows()).toHaveLength(1);

    await publish(item.id, item.token, picture(fakePng(800, 400)));
    click("refresh-links");
    await settle();
    expect(toastText()).toBe("1 linked object.");
    expect(linkRows()[0]?.querySelector(".link-status")?.textContent).toBe(
      "Update available",
    );

    tickRow(0);
    click("update-selected");
    await settle();
    expect(toastText()).toBe("1 updated");
    expect(isError()).toBe(false);
  });

  it("asks for a tick before update selected, go to slide and break", async () => {
    await insertOne();
    for (const id of ["update-selected", "go-to-slide", "break-selected"]) {
      click(id);
      await settle();
      expect(toastText()).toBe("Tick a link in the list first.");
    }
  });

  it("updates the active slide, and says so when the selection is elsewhere", async () => {
    const item = await insertOne();
    await publish(item.id, item.token, picture(fakePng(800, 400)));
    click("refresh-links");
    await settle();

    helpers.selectSlide(presentation.slides[1]!.id);
    click("update-slide");
    await settle();
    expect(toastText()).toBe("No links on this slide");

    helpers.clearSelection();
    click("update-slide");
    await settle();
    expect(toastText()).toBe("Select a slide first.");

    helpers.selectSlide(presentation.slides[0]!.id);
    click("update-slide");
    await settle();
    expect(toastText()).toBe("1 updated");
  });

  it("goes to the slide of the first ticked link and breaks it", async () => {
    await insertOne();
    presentation.moveShape(
      presentation.slides[0]!.shapes[0]!.id,
      presentation.slides[2]!.id,
    );
    click("refresh-links");
    await settle();

    tickRow(0);
    click("go-to-slide");
    await settle();
    expect(toastText()).toBe("Slide 3.");
    expect(presentation.selectedSlideIds).toEqual([presentation.slides[2]!.id]);

    click("break-selected");
    await settle();
    expect(toastText()).toBe("1 link broken. The picture stays on the slide.");
    expect(linkRows()).toHaveLength(0);
  });

  it("keeps a ticked row that a filter hides, and acts on it anyway", async () => {
    const first = await insertOne();
    helpers.selectSlide(presentation.slides[1]!.id);
    await insertOne(await seed(fakePng(400, 200), "Other.xlsx"));
    click("refresh-links");
    await settle();
    expect(linkRows()).toHaveLength(2);

    tickRow(0);
    const filter = document.getElementById(
      "link-source-filter",
    ) as HTMLSelectElement;
    expect([...filter.options].map((option) => option.value)).toEqual([
      "all",
      "Model_v4.xlsx",
      "Other.xlsx",
    ]);
    filter.value = "Other.xlsx";
    filter.dispatchEvent(new Event("change"));
    expect(linkRows()).toHaveLength(1);
    expect(document.getElementById("links-filtered-empty")?.hidden).toBe(true);

    await publish(first.id, first.token, picture(fakePng(800, 400)));
    click("refresh-links");
    await settle();
    click("update-selected");
    await settle();
    // The hidden tick is still a tick: the Map's rule, said out loud.
    expect(toastText()).toBe("1 updated");
  });

  // The list on screen is a snapshot of the last poll, and Excel pushes while
  // this pane sits open: an update that trusts that snapshot reports a deck
  // that has just been re-exported as up to date and repaints nothing.
  it("finds a push the last poll never saw", async () => {
    const item = await insertOne();
    await publish(item.id, item.token, picture(fakePng(1600, 800)));

    click("update-all");
    await settle();

    expect(toastText()).toBe("1 updated");
    expect(presentation.slides[0]!.shapes[0]!.fillImage).toBe(
      fakePng(1600, 800),
    );
  });

  it("says when a filter hides every row", async () => {
    await insertOne();
    click("refresh-links");
    await settle();
    const status = document.getElementById(
      "link-status-filter",
    ) as HTMLSelectElement;
    status.value = "missing";
    status.dispatchEvent(new Event("change"));

    expect(linkRows()).toHaveLength(0);
    expect(document.getElementById("links-filtered-empty")?.hidden).toBe(false);
    expect(document.getElementById("links-empty")?.hidden).toBe(true);
  });
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

    field.value = "not a link key";
    click("save-key");
    await settle();
    expect(toastText()).toBe(
      "That is not a link key. Copy it again from Excel.",
    );

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
});
