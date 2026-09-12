// Shared harness for the stress.ppt.*.integration suites that drive the real
// PowerPoint pane: the shipped pptpane.html in jsdom over the fake host and
// the fake relay, plus the deck fixtures (a planted link, a waiting export)
// and the DOM reads a test makes. Owns no assertions.
// Invariant: every boot is a fresh module graph, deck, relay and workspace.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { vi } from "vitest";
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
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";

export const SRC = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

let relay: FakeRelay;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let workspace: Workspace;

// The pane builds its own RelayClient from document.baseURI; each suite's
// vi.mock factory hands it this one instead.
export function paneRelay(): FakeRelay {
  return relay;
}
export function deck(): FakePresentation {
  return presentation;
}
export function hostHelpers(): FakePptHelpers {
  return helpers;
}
export function paneWorkspace(): Workspace {
  return workspace;
}

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

export function picture(png: string): Payload {
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

export interface Planted {
  shapeId: string;
  id: string;
  token: string;
}

// A link already on slide 2, the way a colleague's deck arrives.
export async function plant(): Promise<Planted> {
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
  return { shapeId: shape.id, id, token };
}

// Excel pushing that link again: the deck's copy is a revision behind.
export async function pushPlanted(planted: Planted): Promise<void> {
  await publish(planted.id, planted.token, picture(fakePng(400, 200)));
}

// One export waiting in the Inbox for this workspace.
export async function waiting(): Promise<InboxItem> {
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

// A fresh pane: new modules, new deck of three slides with the first
// selected, new relay, and the link key already paired.
export async function bootPane(): Promise<void> {
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

export function button(id: string): HTMLButtonElement {
  return document.getElementById(id) as HTMLButtonElement;
}
export function click(id: string): void {
  button(id).click();
}
export async function settle(rounds = 12): Promise<void> {
  for (let round = 0; round < rounds; round += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}
export function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}
// An error toast keeps its reason behind a "Copy details" button.
export function hasDetails(): boolean {
  return document.querySelector("#toast .toast-copy") !== null;
}

// What that button would put on the clipboard: the only way to read the
// details a modeller would paste into a bug report.
export async function copiedDetails(): Promise<string> {
  let copied = "";
  Object.defineProperty(navigator, "clipboard", {
    configurable: true,
    value: {
      writeText: (text: string) => {
        copied = text;
        return Promise.resolve();
      },
    },
  });
  document.querySelector<HTMLButtonElement>("#toast .toast-copy")?.click();
  await settle(2);
  return copied;
}
export function linkRows(): HTMLTableRowElement[] {
  return [...document.querySelectorAll<HTMLTableRowElement>("#link-rows tr")];
}
export function tickRow(index: number): void {
  const box = linkRows()[index]?.querySelector("input");
  if (!box) throw new Error(`no row ${String(index)} to tick`);
  box.checked = true;
  box.dispatchEvent(new Event("change"));
}
export function pressEnterOnKeyField(): void {
  const field = document.getElementById("workspace-key") as HTMLInputElement;
  field.value = workspace.exportKey;
  field.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
  );
}

// A relay call the test releases by hand, so a batch can be caught mid-flight.
export function heldFetch(): () => void {
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
