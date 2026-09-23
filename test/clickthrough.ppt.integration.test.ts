// @vitest-environment jsdom
// Slice H click-through harness (PowerPoint pane): presses every button of
// the shipped pptpane.html + src/ppt/main.ts over the fake PowerPoint host,
// relay mocked, in five shapes - (a) unpaired nothing selected, (b) paired
// nothing selected, (c) paired with two then three shapes selected (object
// tools), (d) paired with an inbox export inserted and its row ticked, (e)
// local mode, a bundle pasted from Excel - and asserts every press answers
// like a product should: no crash, some reaction, never a raw office.js
// code or "undefined" standing alone in the toast, and the busy latch
// releasing again. Also runs every PLSFIX_PPT_* ribbon command, the search
// box and the filter selects.
// Owns no product logic; a real defect found along the way is catalogued in
// KNOWN_DEFECTS below rather than patched here (src/ is out of scope for
// this slice) and the suite asserts it is STILL there, so a later fix must
// remove the entry.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import type * as RelayModule from "../src/link/relay";
import { BUNDLE_SENTENCE, bundleHtml, encodeBundle } from "../src/link/bundle";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import { localWorkspace } from "../src/link/local";
import { LocalCollector } from "../src/link/local-collector";
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
import { TRANSPORT_STORAGE_KEY } from "../src/link/transport-setting";
import {
  createWorkspace,
  WORKSPACE_STORAGE_KEY,
  type Workspace,
} from "../src/link/workspace";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
  type FakeSlide,
} from "./fakeppt";
import { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import { settlePpt, trackPptBoot } from "./ppt-ready";

enableStrictLoadSemantics();

// The pane builds its own RelayClient from document.baseURI; handed this one
// instead, the same technique test/stress.ppt.support.ts uses.
let relay: FakeRelay;
vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return relay;
    },
  };
});

// ---------------------------------------------------------------------------
// Boot. A fresh module graph, deck, relay and workspace every time; paired
// is decided per call rather than reusing test/stress.ppt.support.ts's own
// bootPane() (always pairs) since state (a) needs an unpaired pane too, and
// that file's module-level relay/presentation/workspace belong to its own
// suite's boot, not this one's.
// ---------------------------------------------------------------------------

const SRC = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

let presentation: FakePresentation;
let helpers: FakePptHelpers;
let workspace: Workspace | null;

function pane(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
}

// Waits for the work, not a count of turns: test/ppt-ready.ts's settlePpt
// already drains the fixed rounds then keeps draining while every tab
// button is disabled (the blanket busy latch), bounded, so a press whose
// chain runs long under load is waited out instead of misread as a latch
// that never released.
async function settle(): Promise<void> {
  await settlePpt();
}

interface BootOptions {
  slides?: number;
  paired?: boolean;
  // State (e): local mode. Overrides `paired` - local mode is always its
  // own kind of paired (the fixed local workspace), never the relay's.
  local?: boolean;
}

async function bootPpt(options: BootOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  pane();
  relay = new FakeRelay();
  const storage = new Map<string, string>();
  if (options.local) {
    workspace = null;
    storage.set(TRANSPORT_STORAGE_KEY, "local");
  } else if (options.paired ?? true) {
    workspace = await createWorkspace({
      get: async () => null,
      set: async () => undefined,
      remove: async () => undefined,
    });
    storage.set(WORKSPACE_STORAGE_KEY, workspace.exportKey);
  } else {
    workspace = null;
    // States (a)-(d) are a pairing matrix (relay mode); an unpaired boot
    // would otherwise default to local mode, where everything below is paired
    // by construction and every "unpaired" assertion here would not hold.
    storage.set(TRANSPORT_STORAGE_KEY, "relay");
  }
  const host = installFakePpt({ slides: options.slides ?? 3, storage });
  presentation = host.presentation;
  helpers = host.helpers;
  helpers.selectSlide(presentation.slides[0]!.id);
  const booted = trackPptBoot();
  await import("../src/ppt/main");
  await booted;
  await settle();
}

function findButton(id: string): HTMLButtonElement | null {
  return document.getElementById(id) as HTMLButtonElement | null;
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

// change-source-confirm/cancel live inside #change-source-chooser, which the
// pane keeps `hidden` until "Change source" itself is pressed - see
// scripts/ux/sweep-probe.js's own `hiddenOutsideTab` for the identical rule
// on the Excel side.
vi.stubGlobal("open", vi.fn());
afterEach(() => {
  uninstallFakePpt();
  vi.restoreAllMocks();
  vi.stubGlobal("open", vi.fn());
});

// ---------------------------------------------------------------------------
// Uncaught error / rejection tracking.
// ---------------------------------------------------------------------------

const capturedErrors: string[] = [];
window.addEventListener("error", (event) => {
  capturedErrors.push(`error: ${String(event.error ?? event.message)}`);
});
window.addEventListener("unhandledrejection", (event) => {
  capturedErrors.push(`unhandledrejection: ${String(event.reason)}`);
});

// ---------------------------------------------------------------------------
// Reaction fingerprint - see test/clickthrough.excel.integration.test.ts's
// own note for why the sentinel is spliced directly into the toast's text
// node rather than produced through toast.show().
// ---------------------------------------------------------------------------

const SENTINEL = "CLICKTHROUGH-NO-TOAST-YET";

function armSentinel(): void {
  const container = document.getElementById("toast");
  if (!container) return;
  let span = container.querySelector<HTMLElement>(".toast-text");
  if (!span) {
    span = document.createElement("span");
    span.className = "toast-text";
    container.append(span);
  }
  span.textContent = SENTINEL;
}

// ---------------------------------------------------------------------------
// Busy latch: src/ppt/main.ts's setBusy(busy) disables every ".app-shell
// button" - broader than the Excel pane's [data-action]-only scope - so
// every id-wired PowerPoint button is covered by one shared check. The one
// exception the brief names: change-source, disabled purely by "exactly one
// row ticked" (installChangeSource's own rule), never by busy.
// ---------------------------------------------------------------------------

const BUSY_LATCH_EXCEPTIONS = new Set(["change-source"]);

function enabledSnapshot(): Map<string, boolean> {
  const map = new Map<string, boolean>();
  for (const button of document.querySelectorAll<HTMLButtonElement>(
    ".app-shell button[id]",
  )) {
    if (BUSY_LATCH_EXCEPTIONS.has(button.id)) continue;
    map.set(button.id, !button.disabled);
  }
  return map;
}

function assertBusyReleased(
  pressedKey: string,
  before: Map<string, boolean>,
  after: Map<string, boolean>,
): void {
  for (const [key, wasEnabled] of before) {
    if (!wasEnabled) continue;
    expect(
      after.get(key),
      `after pressing "${pressedKey}", button "${key}" is still disabled (busy latch not released)`,
    ).toBe(true);
  }
}

// ---------------------------------------------------------------------------
// Toast sanity.
// ---------------------------------------------------------------------------

const BARE_HOST_CODE =
  /^(GeneralException|InvalidArgument|ItemAlreadyExists|ItemNotFound|InvalidOperation|InvalidSelection|UnsupportedOperation|AccessDenied|PropertyNotLoaded)$/;

function assertToastSanity(key: string, toast: string): void {
  expect(
    toast,
    `"${key}" toast fell through to an unnamed action fallback`,
  ).not.toMatch(/^Unknown action/);
  expect(
    toast,
    `"${key}" toast is a bare office.js code with no sentence around it`,
  ).not.toMatch(BARE_HOST_CODE);
  expect(
    toast,
    `"${key}" toast leaked the literal text "undefined"`,
  ).not.toMatch(/undefined/);
  expect(toast, `"${key}" toast leaked a stringified object`).not.toMatch(
    /\[object/,
  );
}

// ---------------------------------------------------------------------------
// KNOWN_DEFECTS.
// ---------------------------------------------------------------------------

interface KnownDefect {
  description: string;
  cause: string;
  stillPresent: (toast: string, reacted: boolean) => boolean;
}

const KNOWN_DEFECTS: Record<string, KnownDefect> = {};

// Two icon buttons whose whole job is re-running a read boot() already ran
// eagerly, with nothing in between to make a second run answer differently -
// correctly a no-op, not a dead button (the same shape as
// test/clickthrough.excel.integration.test.ts's refresh-selection/
// refresh-sheets): refresh-links (Office.onReady's own `bootStep(reloadLinks,
// ...)`, unconditional) and refresh-inbox (the same handler's
// `bootStep(refreshInbox, ...)`, run whenever this boot paired).
const IDEMPOTENT_UNTIL_DECK_MOVES = new Set(["refresh-links", "refresh-inbox"]);

// ---------------------------------------------------------------------------
// The one generic press.
// ---------------------------------------------------------------------------

async function pressElement(
  button: HTMLButtonElement,
  label: string,
): Promise<string> {
  if (button.disabled) {
    expect(
      () => button.click(),
      `"${label}" threw synchronously on click`,
    ).not.toThrow();
    await settle();
    return toastText();
  }
  const alreadyActiveTab =
    button.getAttribute("role") === "tab" &&
    button.getAttribute("aria-selected") === "true";
  const expectNoNewReaction =
    alreadyActiveTab || IDEMPOTENT_UNTIL_DECK_MOVES.has(label);
  const busyBefore = enabledSnapshot();
  armSentinel();
  const htmlBefore = document.body.innerHTML;
  const activeBefore = document.activeElement;
  const errorsBefore = capturedErrors.length;

  expect(
    () => button.click(),
    `"${label}" threw synchronously on click`,
  ).not.toThrow();
  await settle();

  expect(
    capturedErrors.slice(errorsBefore),
    `"${label}" raised an uncaught error or unhandled rejection`,
  ).toEqual([]);

  const toast = toastText();
  const htmlAfter = document.body.innerHTML;
  const activeAfter = document.activeElement;
  const reacted =
    toast !== SENTINEL ||
    htmlAfter !== htmlBefore ||
    activeAfter !== activeBefore;

  const defect = KNOWN_DEFECTS[label];
  if (defect) {
    expect(
      defect.stillPresent(toast, reacted),
      `KNOWN_DEFECTS["${label}"] (${defect.description}) no longer reproduces - remove the entry`,
    ).toBe(true);
    return toast;
  }

  expect(
    reacted || expectNoNewReaction,
    `button "${label}" produced no reaction at all (dead button)`,
  ).toBe(true);
  assertToastSanity(label, toast);
  assertBusyReleased(label, busyBefore, enabledSnapshot());
  return toast;
}

async function press(id: string): Promise<string> {
  const button = findButton(id);
  expect(button, `coverage: no button found for "${id}"`).toBeTruthy();
  return pressElement(button!, id);
}

// ---------------------------------------------------------------------------
// Static coverage: every <button> the shipped pptpane.html carries. Every
// one is id-wired; none carries data-action (src/ppt/main.ts's
// BUTTON_ACTIONS table and installChangeSource wire by id alone).
// ---------------------------------------------------------------------------

function staticButtonIds(): string[] {
  const html = readFileSync(join(process.cwd(), "pptpane.html"), "utf8");
  const scratch = new DOMParser().parseFromString(html, "text/html");
  const ids: string[] = [];
  for (const button of scratch.querySelectorAll("button")) {
    if (!button.id) {
      throw new Error(
        `coverage: a pptpane.html button has no id: ${button.outerHTML.slice(0, 80)}`,
      );
    }
    ids.push(button.id);
  }
  return ids;
}

const pressedStatic = new Set<string>();

async function pressStatic(id: string): Promise<string> {
  pressedStatic.add(id);
  return press(id);
}

type GeneratedKind = "insert-button" | "link-row-tick";
const generatedCovered = new Set<GeneratedKind>();

// ---------------------------------------------------------------------------
// Seeding.
// ---------------------------------------------------------------------------

function picturePayload(png: string): Payload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 200,
    height: 100,
    png,
    src: SRC,
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
}

function newId(): string {
  return newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
}

async function publishPicture(
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

// One export waiting in the Inbox for the currently paired workspace.
async function seedWaitingInbox(): Promise<InboxItem> {
  const ws = workspace;
  if (!ws) throw new Error("seed: seedWaitingInbox needs a paired boot");
  const id = newId();
  const token = newToken();
  await publishPicture(id, token, picturePayload(fakePng(200, 100)));
  const item: InboxItem = {
    id,
    token,
    kind: "range",
    label: "Model!B4:F12",
    src: SRC,
    createdAt: new Date().toISOString(),
  };
  await relay.postInbox(
    ws.id,
    ws.auth,
    id,
    await seal(ws.enc, ws.id, encodeInboxItem(item)),
  );
  return item;
}

// ---------------------------------------------------------------------------
// Local mode: a synthetic clipboard paste, built the way Excel's
// LocalCollector records an export (src/link/local-collector.ts) and carried
// the way a real copy is - the HTML flavor first, text/plain the sentence.
// ---------------------------------------------------------------------------

function pasteEvent(html: string, plain: string): Event {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => (type === "text/html" ? html : plain) },
  });
  return event;
}

function bundleEvent(bundle: ReturnType<LocalCollector["bundle"]>): Event {
  return pasteEvent(bundleHtml(encodeBundle(bundle)), BUNDLE_SENTENCE);
}

function pasteBox(): HTMLTextAreaElement {
  return document.getElementById("paste-links") as HTMLTextAreaElement;
}

// One Excel export the way LocalCollector records it: the link's sealed
// payload plus its inbox row (design 2026-09-23 s.4 - every export copies
// both; LocalStore.ingest is what decides the row is kept or dropped).
async function exportLinkLocally(
  collector: LocalCollector,
  id: string,
  token: string,
  payload: Payload,
  currentRev = 0,
): Promise<void> {
  const keys = await deriveLinkKeys(token);
  await collector.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
    currentRev,
  );
  const localWs = await localWorkspace();
  const item: InboxItem = {
    id,
    token,
    kind: "range",
    label: "Model!B4:F12",
    src: SRC,
    createdAt: new Date().toISOString(),
  };
  await collector.postInbox(
    localWs.id,
    localWs.auth,
    id,
    await seal(localWs.enc, localWs.id, encodeInboxItem(item)),
  );
}

// States (b)/(c)/(d): "paired" the relay way (a real key) or the local way
// (the fixed local workspace, bootPpt's own `local: true`) - the two ways
// src/ppt/transport.ts answers workspace() for every deck flow below.
async function bootPaired(transport: "relay" | "local"): Promise<void> {
  await bootPpt(transport === "relay" ? { paired: true } : { local: true });
}

// One export left waiting in the Inbox, and the pane's own view of it
// refreshed - relay's seedWaitingInbox against the paired workspace, or a
// fresh export pasted as a bundle against the local store (ingest() parks an
// id the deck does not hold yet in the Inbox exactly like a relay push does,
// so both leave one ".inbox-insert" ready to click).
async function seedWaitingExport(transport: "relay" | "local"): Promise<void> {
  if (transport === "relay") {
    await seedWaitingInbox();
    await press("refresh-inbox");
    return;
  }
  const collector = new LocalCollector();
  await exportLinkLocally(
    collector,
    newId(),
    newToken(),
    picturePayload(fakePng(200, 100)),
  );
  pasteBox().dispatchEvent(bundleEvent(collector.bundle()));
  await settle();
}

// A shape already on the given slide, tagged at `rev`: the deck a bundle is
// about to update, the way a colleague's deck arrives.
function plantLink(
  slide: FakeSlide,
  id: string,
  token: string,
  rev: number,
): void {
  const shape = presentation.addShape(slide, {
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
    rev,
    src: SRC,
    pushedAt: new Date().toISOString(),
  };
  shape.tags.set(TAG_LINK, encodeTag(tag));
  shape.tags.set(TAG_KEY, token);
}

// Three same-size, same-type shapes on the given slide, so
// sameKindAndSize (src/ppt/object-math.ts) matches every pair.
function addPlainShapes(slide: FakeSlide, count: number): FakePptShape[] {
  const shapes: FakePptShape[] = [];
  for (let i = 0; i < count; i += 1) {
    shapes.push(
      presentation.addShape(slide, {
        type: "GeometricShape",
        geometry: "Rectangle",
        left: 40 + i * 140,
        top: 60,
        width: 100,
        height: 60,
      }),
    );
  }
  return shapes;
}

function linkRows(): HTMLTableRowElement[] {
  return [
    ...document.querySelectorAll<HTMLTableRowElement>(
      "#link-rows tr[data-key]",
    ),
  ];
}

function tickRow(index: number): void {
  const box = linkRows()[index]?.querySelector("input");
  if (!box) throw new Error(`no row ${String(index)} to tick`);
  box.checked = true;
  box.dispatchEvent(new Event("change"));
}

// ---------------------------------------------------------------------------
// State (a): nothing selected, unpaired (storage empty).
// ---------------------------------------------------------------------------

describe("state (a): unpaired, nothing selected", () => {
  it("presses every static button and every ribbon command once", async () => {
    await bootPpt({ paired: false });

    for (const id of new Set(staticButtonIds())) {
      await pressStatic(id);
    }

    for (const id of helpers.commandIds()) {
      await helpers.runCommand(id);
      await settle();
    }
  });

  it("answers the brief's specific sentences unpaired", async () => {
    await bootPpt({ paired: false });

    for (const id of ["update-selected", "go-to-slide"]) {
      expect(await press(id)).toBe("Tick a link in the list first.");
    }
    // break-selected only arms on the first press (src/ui/confirm.ts); the
    // second is what actually runs and meets the empty selection.
    await press("break-selected");
    expect(await press("break-selected")).toBe(
      "Tick a link in the list first.",
    );
    // revertSelected() passes its own message to requireSelection()
    // (src/ppt/main.ts), the one PPT-side button with a rule of its own
    // here, the way copy-model-check/trace-back etc. do on the Excel side.
    expect(await press("revert-selected")).toBe("Tick the rows to revert.");
    expect(await press("refresh-inbox")).toBe("Paste the link key in Settings");
    expect(await press("paste-latest-linked")).toBe(
      "Paste the link key in Settings",
    );
    expect(await press("save-key")).toBe(
      "Paste the link key from Excel first.",
    );
    expect(
      findButton("change-source")!.disabled,
      "0 ticked rows: change-source stays disabled",
    ).toBe(true);
  });
});

// States (b)-(d) below run once per transport: relay (a real key, the
// pairing matrix's own paired half) then local (the fixed local
// workspace, bootPaired's own local branch, no key stored). State (a)
// stays relay-only ("unpaired" has no local analog: local mode is always
// paired by construction) and state (e) stays its own dedicated pass
// (paste mechanics have no relay analog).
describe.each(["relay", "local"] as const)("transport: %s", (transport) => {
  // ---------------------------------------------------------------------------
  // State (b): paired, nothing selected.
  // ---------------------------------------------------------------------------

  describe("state (b): paired, nothing selected", () => {
    it("presses every static button and every ribbon command once", async () => {
      await bootPaired(transport);

      for (const id of new Set(staticButtonIds())) {
        await pressStatic(id);
      }

      for (const id of helpers.commandIds()) {
        await helpers.runCommand(id);
        await settle();
      }
    });

    it("filters the link list by typing, and every filter select reacts once", async () => {
      await bootPaired(transport);
      // A row to filter: state (b) itself seeds no deck links, and an empty
      // list hides the "no match" message regardless of query
      // (renderLinks(): linksFilteredEmpty.hidden = rows.length === 0 || ...).
      await seedWaitingExport(transport);
      document.querySelector<HTMLButtonElement>(".inbox-insert")!.click();
      await settle();
      expect(linkRows().length, "seed: expected one inserted link row").toBe(1);

      const search = document.getElementById("link-search") as HTMLInputElement;
      const before = document.body.innerHTML;
      search.value = "nothing matches this query";
      search.dispatchEvent(new Event("input"));
      await settle();
      expect(
        document.body.innerHTML,
        "typing into the search box changed nothing",
      ).not.toBe(before);
      expect(document.getElementById("links-filtered-empty")?.hidden).toBe(
        false,
      );
      search.value = "";
      search.dispatchEvent(new Event("input"));
      await settle();

      for (const id of [
        "link-status-filter",
        "link-source-filter",
        "link-slide-filter",
        "link-project-filter",
      ]) {
        const select = document.getElementById(id) as HTMLSelectElement;
        for (let i = 0; i < select.options.length; i += 1) {
          select.selectedIndex = i;
          expect(() => select.dispatchEvent(new Event("change"))).not.toThrow();
          await settle();
        }
      }
    });
  });

  // ---------------------------------------------------------------------------
  // State (c): paired, object tools over two then three same-size shapes.
  // ---------------------------------------------------------------------------

  describe("state (c): paired, two then three shapes selected", () => {
    it("aligns, matches and swaps two, then distributes three", async () => {
      await bootPaired(transport);
      const slide = presentation.slides[0]!;
      const shapes = addPlainShapes(slide, 3);

      helpers.selectShapes([shapes[0]!.id, shapes[1]!.id]);
      expect(await pressStatic("align-objects")).toBe(
        "Aligned 2 objects left.",
      );
      expect(await pressStatic("match-size")).toBe(
        "Matched 1 object to the reference size.",
      );
      expect(await pressStatic("swap-objects")).toBe("Swapped two objects.");

      helpers.selectShapes(shapes.map((shape) => shape.id));
      expect(await pressStatic("distribute-objects")).toBe(
        "Distributed 3 objects across.",
      );

      helpers.selectShapes([shapes[0]!.id]);
      expect(await pressStatic("select-similar")).toBe(
        "Selected 3 similar objects.",
      );

      helpers.selectShapes([shapes[0]!.id]);
      expect(await pressStatic("capture-object-style")).toBe(
        "Object style captured.",
      );
      helpers.selectShapes([shapes[1]!.id, shapes[2]!.id]);
      expect(await pressStatic("apply-object-style")).toBe(
        "Painted 2 objects.",
      );
    });

    it("runs every PLSFIX_PPT_* ribbon command over a real selection", async () => {
      await bootPaired(transport);
      const slide = presentation.slides[0]!;
      const shapes = addPlainShapes(slide, 3);
      helpers.selectShapes(shapes.map((shape) => shape.id));

      for (const id of helpers.commandIds()) {
        const before = toastText();
        await helpers.runCommand(id);
        await settle();
        assertToastSanity(id, toastText());
        // PLSFIX_PPT_TOOLS answers with "" (it only opens the pane); every
        // other command over three consistent shapes has real work to do.
        if (id !== "PLSFIX_PPT_TOOLS") {
          expect(toastText(), `ribbon "${id}" produced no new toast`).not.toBe(
            before,
          );
        }
        helpers.selectShapes(shapes.map((shape) => shape.id));
      }
    });
  });

  // ---------------------------------------------------------------------------
  // State (d): paired, an inbox export inserted, its row ticked - and a
  // second waiting export, so "Change source" has a real candidate.
  // ---------------------------------------------------------------------------

  describe("state (d): paired, an inserted link, its row ticked", () => {
    it("updates, reverts, jumps, re-points and breaks a real linked row", async () => {
      await bootPaired(transport);
      await seedWaitingExport(transport);

      const insertButton =
        document.querySelector<HTMLButtonElement>(".inbox-insert");
      expect(
        insertButton,
        "seed: expected an Insert button in the inbox",
      ).toBeTruthy();
      generatedCovered.add("insert-button");
      await pressElement(insertButton!, "inbox-insert:first");

      // A second export, left waiting, so change-source has something to
      // re-point at.
      await seedWaitingExport(transport);

      expect(linkRows().length, "seed: expected one inserted link row").toBe(1);
      tickRow(0);
      generatedCovered.add("link-row-tick");

      expect(await pressStatic("update-selected")).toMatch(
        /up to date|updated/i,
      );
      expect(await pressStatic("update-slide")).toMatch(
        /up to date|updated|Select a slide first\.|No links on this slide/,
      );
      expect(await pressStatic("update-all")).toMatch(/up to date|updated/i);
      expect(await pressStatic("go-to-slide")).toMatch(/^Slide \d+\.$/);

      tickRow(0);
      expect(
        findButton("change-source")!.disabled,
        "exactly one row ticked: enabled",
      ).toBe(false);
      expect(await pressStatic("change-source")).toMatch(
        /^Choose the export to point/,
      );
      expect(await pressStatic("change-source-cancel")).toBe(
        "Change source cancelled. Nothing was re-pointed.",
      );
      await pressStatic("change-source");
      // Both seeded items share SRC, so changeSource()'s own line
      // (src/ppt/change-source.ts) names the same workbook on both sides.
      expect(await pressStatic("change-source-confirm")).toMatch(
        /^Source changed: Model_v4\.xlsx -> Model_v4\.xlsx$/,
      );

      tickRow(0);
      // summarizeRevert() (src/ppt/revert.ts): "N reverted" when a previous
      // revision exists, "N without a previous version" otherwise - the
      // change-source press above put this row on its very first revision,
      // so either wording is a legitimate answer; a bare host string or an
      // empty toast is not.
      expect(await pressStatic("revert-selected")).toMatch(
        /^\d+ (reverted|without a previous version|failed)|^Nothing to revert$/,
      );
      tickRow(0);
      // break-selected only arms on the first press (src/ui/confirm.ts); the
      // second is what actually breaks it.
      await pressStatic("break-selected");
      expect(await pressStatic("break-selected")).toMatch(/^1 link broken\./);
    });
  });
});

// ---------------------------------------------------------------------------
// State (e): local mode - no key stored at all, transport pinned local. One
// paste can both repaint a link already on a slide and park a brand new one
// in the Inbox; a copy that is not a pls,fix bundle says so in one sentence.
// ---------------------------------------------------------------------------

describe("state (e): local mode, a bundle pasted from Excel", () => {
  it("presses every static button and every ribbon command once", async () => {
    await bootPpt({ local: true });

    for (const id of new Set(staticButtonIds())) {
      await pressStatic(id);
    }

    for (const id of helpers.commandIds()) {
      await helpers.runCommand(id);
      await settle();
    }
  });

  it("repaints a held link and parks a new one in the Inbox, from one paste", async () => {
    await bootPpt({ local: true, slides: 3 });
    const heldId = newId();
    const heldToken = newToken();
    plantLink(presentation.slides[1]!, heldId, heldToken, 1);
    const freshId = newId();
    const freshToken = newToken();

    const collector = new LocalCollector();
    await exportLinkLocally(
      collector,
      heldId,
      heldToken,
      picturePayload(fakePng(400, 200)),
      1,
    );
    await exportLinkLocally(
      collector,
      freshId,
      freshToken,
      picturePayload(fakePng(200, 100)),
    );

    pasteBox().dispatchEvent(bundleEvent(collector.bundle()));
    await settle();

    expect(toastText()).toBe(
      "Pasted 2 links: 1 updated, 1 waiting in the Inbox.",
    );
    expect(presentation.slides[1]!.shapes[0]!.fillImage).toBe(
      fakePng(400, 200),
    );
    // inboxQuietly() inside the paste itself re-reads the Inbox, so the new
    // link's Insert button is already there with no extra press.
    expect(
      document.querySelector<HTMLButtonElement>(".inbox-insert"),
    ).toBeTruthy();
  });

  it("says one plain sentence for a paste that is not a pls,fix copy", async () => {
    await bootPpt({ local: true });

    pasteBox().dispatchEvent(pasteEvent("", "not a bundle at all"));
    await settle();

    expect(toastText()).toBe("That is not a pls,fix copy from Excel.");
  });

  it("never fills the box: typing is wiped back out", async () => {
    await bootPpt({ local: true });
    const box = pasteBox();

    box.value = "something typed by hand";
    box.dispatchEvent(new Event("input", { bubbles: true }));

    expect(box.value).toBe("");
  });
});

// ---------------------------------------------------------------------------
// Two-click buttons: the first press only arms; the second is the real one.
// ---------------------------------------------------------------------------

describe("two-click confirms", () => {
  it("break-selected arms on the first click and breaks on the second", async () => {
    await bootPpt({ paired: true });
    await seedWaitingInbox();
    await pressStatic("refresh-inbox");
    document.querySelector<HTMLButtonElement>(".inbox-insert")!.click();
    await settle();
    expect(linkRows().length, "seed: expected one inserted link row").toBe(1);
    tickRow(0);

    const button = findButton("break-selected")!;
    expect(button.classList.contains("armed")).toBe(false);
    await press("break-selected");
    expect(
      button.classList.contains("armed"),
      "first press should only arm it",
    ).toBe(true);

    const secondToast = await press("break-selected");
    expect(secondToast).toBe("1 link broken. The object stays on the slide.");
    expect(
      button.classList.contains("armed"),
      "the confirmed break disarms it again",
    ).toBe(false);
  });

  it("forget-key arms on the first click and forgets on the second", async () => {
    await bootPpt({ paired: true });

    const button = findButton("forget-key")!;
    expect(button.classList.contains("armed")).toBe(false);
    await press("forget-key");
    expect(
      button.classList.contains("armed"),
      "first press should only arm it",
    ).toBe(true);

    const secondToast = await press("forget-key");
    expect(secondToast).toBe(
      "Link key forgotten. The links already in this deck still update.",
    );
    expect(
      button.classList.contains("armed"),
      "the confirmed forget disarms it again",
    ).toBe(false);
  });

  it("clear-pasted-links arms on the first click and clears on the second", async () => {
    await bootPpt({ local: true });
    const collector = new LocalCollector();
    await exportLinkLocally(
      collector,
      newId(),
      newToken(),
      picturePayload(fakePng(200, 100)),
    );
    pasteBox().dispatchEvent(bundleEvent(collector.bundle()));
    await settle();
    expect(
      document.querySelectorAll("#inbox-list button").length,
    ).toBeGreaterThan(0);

    const button = findButton("clear-pasted-links")!;
    expect(button.classList.contains("armed")).toBe(false);
    await press("clear-pasted-links");
    expect(
      button.classList.contains("armed"),
      "first press should only arm it",
    ).toBe(true);

    const secondToast = await press("clear-pasted-links");
    expect(secondToast).toBe(
      "Pasted links cleared from this computer. The deck keeps its objects.",
    );
    expect(
      button.classList.contains("armed"),
      "the confirmed clear disarms it again",
    ).toBe(false);

    await pressStatic("refresh-inbox");
    expect(toastText()).toBe("Nothing waiting from Excel.");
  });
});

// ---------------------------------------------------------------------------
// Coverage: every static button the shipped pptpane.html carries, and every
// documented generated-content kind, was pressed by at least one state/pass
// above.
// ---------------------------------------------------------------------------

describe("coverage", () => {
  it("pressed every static button at least once", () => {
    const required = new Set(staticButtonIds());
    const missing = [...required].filter((id) => !pressedStatic.has(id));
    expect(missing, "static buttons never pressed by any state").toEqual([]);
  });

  it("exercised every documented generated-content kind at least once", () => {
    const required: GeneratedKind[] = ["insert-button", "link-row-tick"];
    const missing = required.filter((kind) => !generatedCovered.has(kind));
    expect(missing, "generated-content kinds never exercised").toEqual([]);
  });
});
