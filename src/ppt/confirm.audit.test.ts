// @vitest-environment jsdom
// The PowerPoint pane's two confirm buttons (break-selected, forget-key):
// the first press only arms and lapses on its own, src/ui/confirm.ts owns
// the rest. Split out of main.audit.test.ts, already at the file's own
// 400-line cap, rather than pushed further over it.

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
import { settlePpt, trackPptBoot } from "../../test/ppt-ready";
import { CONFIRM_MS } from "../ui/confirm";
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

// The pane boots on import, so the markup, the deck and the pairing key in
// storage all go in first, one linked object already on the active slide -
// inserted through the pane's own Insert button, the same route
// main.audit.test.ts's insertOne() takes.
async function boot(): Promise<void> {
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
  const storage = new Map([[WORKSPACE_STORAGE_KEY, workspace.exportKey]]);
  const host = installFakePpt({ slides: 3, storage });
  presentation = host.presentation;
  helpers = host.helpers;
  helpers.selectSlide(presentation.slides[0]!.id);
  const booted = trackPptBoot();
  await import("./main");
  await booted;
  await settle();

  await seed();
  click("refresh-inbox");
  await settle();
  document
    .querySelectorAll<HTMLButtonElement>("#inbox-list button")[0]
    ?.click();
  await settle();
  click("refresh-links");
  await settle();
}

function click(id: string): void {
  document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
}

function button(id: string): HTMLButtonElement {
  return document.getElementById(id) as HTMLButtonElement;
}

async function settle(rounds = 12): Promise<void> {
  await settlePpt(rounds);
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function linkRows(): HTMLTableRowElement[] {
  const selector = "#link-rows tr[data-key]";
  return [...document.querySelectorAll<HTMLTableRowElement>(selector)];
}

function tickRow(index: number): void {
  const box = linkRows()[index]?.querySelector("input");
  if (!box) throw new Error(`no row ${String(index)} to tick`);
  box.checked = true;
  box.dispatchEvent(new Event("change"));
}

beforeEach(async () => {
  await boot();
});
afterEach(() => {
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("two-click confirms", () => {
  it("break-selected and forget-key arm on the first press and lapse without running", async () => {
    tickRow(0);

    vi.useFakeTimers();
    try {
      for (const id of ["break-selected", "forget-key"]) {
        click(id);
        expect(
          button(id).classList.contains("armed"),
          `${id}: first press should only arm it`,
        ).toBe(true);
        vi.advanceTimersByTime(CONFIRM_MS);
        expect(
          button(id).classList.contains("armed"),
          `${id}: five seconds on, it should have disarmed itself`,
        ).toBe(false);
      }
    } finally {
      vi.useRealTimers();
    }

    // Neither ever ran: the link is intact and the key is still paired.
    expect(linkRows()).toHaveLength(1);
    expect(document.getElementById("workspace-state")?.textContent).toBe(
      "Paired",
    );
  });

  it("breaks the ticked link only on the second press", async () => {
    tickRow(0);

    click("break-selected");
    expect(button("break-selected").classList.contains("armed")).toBe(true);
    expect(linkRows()).toHaveLength(1);

    click("break-selected");
    await settle();
    expect(toastText()).toBe("1 link broken. The object stays on the slide.");
    expect(linkRows()).toHaveLength(0);
  });

  it("forgets the key only on the second press", async () => {
    click("forget-key");
    expect(document.getElementById("workspace-state")?.textContent).toBe(
      "Paired",
    );

    click("forget-key");
    await settle();
    expect(toastText()).toBe(
      "Link key forgotten. The links already in this deck still update.",
    );
    expect(document.getElementById("workspace-state")?.textContent).toContain(
      "Not paired",
    );
  });
});
