// @vitest-environment jsdom
// The PowerPoint pane's "Change source" picker, driven through the real
// pptpane.html markup: the button's one-tick rule, the candidate list, the
// re-point itself, the kind warning and Cancel. The fake deck and the
// in-memory relay stand in for the host and the server.

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

describe("the Change source picker", () => {
  beforeEach(async () => {
    await boot();
  });

  it("opens on one tick, re-points the picture and closes", async () => {
    await insertOne();
    const newer = await seed(fakePng(1600, 800), "Model_v9.xlsx");
    click("refresh-inbox");
    await settle();
    expect(button("change-source").disabled).toBe(true);

    tickRow(0);
    expect(button("change-source").disabled).toBe(false);
    click("change-source");
    await settle();
    expect(toastText()).toBe("Choose the export to point Model!B4:F12 at.");
    const list = document.getElementById(
      "change-source-list",
    ) as HTMLSelectElement;
    expect(document.getElementById("change-source-chooser")?.hidden).toBe(
      false,
    );
    expect([...list.options].map((option) => option.value)).toEqual([newer.id]);

    click("change-source-confirm");
    await settle();
    expect(toastText()).toBe("Source changed: Model_v4.xlsx -> Model_v9.xlsx");
    expect(document.getElementById("change-source-chooser")?.hidden).toBe(true);
    expect(presentation.slides[0]!.shapes[0]!.fillImage).toBe(
      fakePng(1600, 800),
    );
    // The export it consumed leaves the Inbox with it.
    expect(inboxButtons()).toHaveLength(0);
  });

  it("says so with nothing waiting, and closes when the tick goes", async () => {
    await insertOne();
    tickRow(0);
    click("change-source");
    await settle();
    expect(toastText()).toBe(
      "Nothing waiting in the Inbox. Export the range again from Excel first.",
    );
    expect(document.getElementById("change-source-chooser")?.hidden).toBe(true);

    await seed(fakePng(1600, 800), "Model_v9.xlsx");
    click("refresh-inbox");
    await settle();
    click("change-source");
    await settle();
    expect(document.getElementById("change-source-chooser")?.hidden).toBe(
      false,
    );

    const box = linkRows()[0]!.querySelector("input")!;
    box.checked = false;
    box.dispatchEvent(new Event("change"));
    expect(button("change-source").disabled).toBe(true);
    expect(document.getElementById("change-source-chooser")?.hidden).toBe(true);
  });

  it("warns in the details when the new source is another kind", async () => {
    await insertOne();
    await seed(fakePng(1600, 800), "Model_v9.xlsx", "chart");
    click("refresh-inbox");
    await settle();
    tickRow(0);
    click("change-source");
    await settle();

    click("change-source-confirm");
    await settle();

    expect(toastText()).toBe("Source changed: Model_v4.xlsx -> Model_v9.xlsx");
    // kindWarning is staged as the toast's details, not swallowed.
    expect(document.querySelector("#toast .toast-copy")).not.toBeNull();
  });

  it("cancels without re-pointing anything", async () => {
    await insertOne();
    await seed(fakePng(1600, 800), "Model_v9.xlsx");
    click("refresh-inbox");
    await settle();
    tickRow(0);
    click("change-source");
    await settle();

    click("change-source-cancel");
    await settle();

    expect(toastText()).toBe(
      "Change source cancelled. Nothing was re-pointed.",
    );
    expect(document.getElementById("change-source-chooser")?.hidden).toBe(true);
    expect(presentation.slides[0]!.shapes[0]!.fillImage).toBe(
      fakePng(800, 400),
    );
  });
});
