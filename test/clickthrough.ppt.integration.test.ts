// @vitest-environment jsdom
// Slice H click-through harness (PowerPoint pane): presses every button of
// the shipped pptpane.html + src/ppt/main.ts over the fake PowerPoint host,
// relay mocked, in four shapes - (a) unpaired nothing selected, (b) paired
// nothing selected, (c) paired with two then three shapes selected (object
// tools), (d) paired with an inbox export inserted and its row ticked - and
// asserts every press answers like a product should: no crash, some
// reaction, never a raw office.js code or "undefined" standing alone in the
// toast, and the busy latch releasing again. Also runs every PLSFIX_PPT_*
// ribbon command, the search box and the filter selects.
// Owns no product logic; a real defect found along the way is catalogued in
// KNOWN_DEFECTS below rather than patched here (src/ is out of scope for
// this slice) and the suite asserts it is STILL there, so a later fix must
// remove the entry.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import type * as RelayModule from "../src/link/relay";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import {
  encodeInboxItem,
  encodePayload,
  newLinkId,
  type InboxItem,
  type Payload,
} from "../src/link/model";
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
}

async function bootPpt(options: BootOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  pane();
  relay = new FakeRelay();
  const storage = new Map<string, string>();
  if (options.paired ?? true) {
    workspace = await createWorkspace({
      get: async () => null,
      set: async () => undefined,
      remove: async () => undefined,
    });
    storage.set(WORKSPACE_STORAGE_KEY, workspace.exportKey);
  } else {
    workspace = null;
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

// ---------------------------------------------------------------------------
// State (b): paired, nothing selected.
// ---------------------------------------------------------------------------

describe("state (b): paired, nothing selected", () => {
  it("presses every static button and every ribbon command once", async () => {
    await bootPpt({ paired: true });

    for (const id of new Set(staticButtonIds())) {
      await pressStatic(id);
    }

    for (const id of helpers.commandIds()) {
      await helpers.runCommand(id);
      await settle();
    }
  });

  it("filters the link list by typing, and every filter select reacts once", async () => {
    await bootPpt({ paired: true });
    // A row to filter: state (b) itself seeds no deck links, and an empty
    // list hides the "no match" message regardless of query
    // (renderLinks(): linksFilteredEmpty.hidden = rows.length === 0 || ...).
    await seedWaitingInbox();
    await press("refresh-inbox");
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
    expect(document.getElementById("links-filtered-empty")?.hidden).toBe(false);
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
    await bootPpt({ paired: true });
    const slide = presentation.slides[0]!;
    const shapes = addPlainShapes(slide, 3);

    helpers.selectShapes([shapes[0]!.id, shapes[1]!.id]);
    expect(await pressStatic("align-objects")).toBe("Aligned 2 objects left.");
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
    expect(await pressStatic("apply-object-style")).toBe("Painted 2 objects.");
  });

  it("runs every PLSFIX_PPT_* ribbon command over a real selection", async () => {
    await bootPpt({ paired: true });
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
    await bootPpt({ paired: true });
    await seedWaitingInbox();
    await pressStatic("refresh-inbox");

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
    await seedWaitingInbox();
    await pressStatic("refresh-inbox");

    expect(linkRows().length, "seed: expected one inserted link row").toBe(1);
    tickRow(0);
    generatedCovered.add("link-row-tick");

    expect(await pressStatic("update-selected")).toMatch(/up to date|updated/i);
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
