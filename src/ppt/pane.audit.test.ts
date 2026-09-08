// @vitest-environment jsdom
// The two states the PowerPoint pane has to survive on its own: a deck with
// no pairing key, where the Links list, Go to slide, Break and Update still
// work and only the relay actions refuse; and a relay that answers nothing,
// where every action fails with one sentence and hands the pane back.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../link/crypto";
import {
  encodePayload,
  encodeTag,
  newLinkId,
  TAG_KEY,
  TAG_LINK,
  type LinkTag,
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

// A link already in the deck, the way a colleague's deck arrives: a tagged
// picture this machine never inserted and holds no workspace key for.
async function plant(): Promise<void> {
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

function tickRow(index: number): void {
  const box = linkRows()[index]?.querySelector("input");
  if (!box) throw new Error(`no row ${String(index)} to tick`);
  box.checked = true;
  box.dispatchEvent(new Event("change"));
}

afterEach(() => {
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("a deck with no link key", () => {
  beforeEach(async () => {
    await boot(false);
    await plant();
    click("refresh-links");
    await settle();
  });

  it("lists the links it holds and updates them", async () => {
    expect(document.getElementById("workspace-state")?.textContent).toContain(
      "Not paired",
    );
    expect(document.getElementById("inbox-list")?.hidden).toBe(true);
    expect(document.getElementById("inbox-unpaired")?.hidden).toBe(false);
    expect(linkRows()).toHaveLength(1);
    expect(toastText()).toBe("1 linked object.");

    click("update-all");
    await settle();
    expect(toastText()).toBe("1 up to date");
  });

  it("goes to a slide and breaks a link without a key", async () => {
    tickRow(0);
    click("go-to-slide");
    await settle();
    expect(toastText()).toBe("Slide 2.");

    click("break-selected");
    await settle();
    expect(toastText()).toBe("1 link broken. The picture stays on the slide.");
    expect(linkRows()).toHaveLength(0);
    expect(presentation.slides[1]!.shapes[0]!.fillImage).toBe(
      fakePng(800, 400),
    );
  });

  it("refuses only the actions that need the key, in one sentence", async () => {
    for (const id of ["refresh-inbox", "paste-latest-linked"]) {
      click(id);
      await settle();
      expect(toastText()).toBe("Paste the link key in Settings");
      expect(button(id).disabled).toBe(false);
    }

    tickRow(0);
    click("change-source");
    await settle();
    expect(toastText()).toBe("Paste the link key in Settings");
    expect(document.getElementById("change-source-chooser")?.hidden).toBe(true);
  });
});

describe("the relay unreachable", () => {
  const down = (): RelayError =>
    new RelayError("network", "relay POST /api/links/status: network error");

  beforeEach(async () => {
    await boot();
    await plant();
    click("refresh-links");
    await settle();
    for (const call of [
      "status",
      "fetchLinks",
      "getLink",
      "getLinkRev",
      "listInbox",
      "deleteInbox",
    ] as const) {
      vi.spyOn(relay, call).mockRejectedValue(down());
    }
  });

  it("fails every relay action with one sentence and leaves the pane usable", async () => {
    for (const id of [
      "refresh-links",
      "refresh-inbox",
      "paste-latest-linked",
    ]) {
      click(id);
      await settle();
      expect(toastText()).toBe("relay POST /api/links/status: network error");
      // The busy flag is off again on every one of them.
      expect(button("refresh-links").disabled).toBe(false);
      expect(button("update-all").disabled).toBe(false);
    }
    // The list the last good poll drew is still there to act on.
    expect(linkRows()).toHaveLength(1);
  });

  it("reports the rows an update and a revert could not reach", async () => {
    click("update-all");
    await settle();
    expect(toastText()).toBe("1 failed");
    expect(button("update-all").disabled).toBe(false);

    tickRow(0);
    click("revert-selected");
    await settle();
    // rev 1 has nothing below it, so the relay is never asked.
    expect(toastText()).toBe("1 without a previous version");
    expect(button("revert-selected").disabled).toBe(false);
  });

  it("clears the busy flag after something that is not an Error", async () => {
    vi.spyOn(relay, "status").mockRejectedValue("relay exploded");

    click("refresh-links");
    await settle();

    expect(toastText()).toBe("The add-in could not complete that action.");
    expect(button("refresh-links").disabled).toBe(false);
    expect(button("break-selected").disabled).toBe(false);
  });
});
