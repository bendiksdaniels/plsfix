# Text Links Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** One Excel cell becomes a tagged, auto-sized text box on a slide that refreshes in place; plus the Links tab tells the relay on boot which links the workbook still holds.

**Architecture:** A third payload kind `text` rides the existing link pipeline unchanged (hidden-name anchor, registry entry, sealed relay blob, inbox note, tags). Excel renders `Range.text[0][0]`; PowerPoint inserts `addTextBox` and refreshes `textFrame.textRange.text` only. Everything else (auto-push, highlight, revert, break, Update all) already works on tags and revisions.

**Tech Stack:** TypeScript, Office.js (ExcelApi 1.9 floor, PowerPointApi 1.4 for text boxes), Vitest with the strict fake hosts in `test/`.

**Spec:** `docs/superpowers/specs/2026-08-31-text-links-design.md`

## Global Constraints

- Every source file at most 400 lines, every function at most 50 lines; each new file opens with a 2-4 line header comment (purpose, what it owns, key invariant). Split at the cap.
- Every Office.js scalar read is preceded by `load()` + `sync()`; the strict fakes throw otherwise.
- Link identity is the tags (`PLSFIX_LINK`, `PLSFIX_KEY`) in PowerPoint and the hidden name `PLSFIX_LINK_<id8>` in Excel. A refresh never writes geometry.
- The server sees ciphertext only; nothing in this plan touches `server/`.
- No `console.*` unless the codebase already does it in that layer (check `eslint.config.js` and `git grep -n "console\." src`).
- `TEXT_MAX_CHARS` = 500. One cell only in v1 (a merged area is refused; the spec's ExcelApi 1.13 merged-area rule is deferred, see Task 7).
- Copy rules: pane strings in English; the Latvian guide and manual paragraphs are given verbatim below; no em dashes anywhere.
- Gates before every commit: `npm run check`; Task 6 also `npm run ux:check` (needs `npm run dev` running). Commit messages: short imperative subject, plain 3-6 line body, no trailers.
- Do not bump the version, touch the manifests or deploy; the controller does that at merge.

---

## File structure

| File | Responsibility |
|---|---|
| `src/link/model.ts` (modify) | `LinkKind` gains `"text"`, `TextPayload`, `TEXT_MAX_CHARS`, the payload guard, `sourceLabel` for text |
| `src/excel/link-touch.ts` (create) | `touchWorkbookLinks(relay)`: registry -> auth keys -> `relay.touchLinks` |
| `src/excel/link-anchors.ts` (modify) | `ResolvedRange.kind` and `Render` gain text; `renderSource` text branch |
| `src/excel/link-record.ts` (modify) | `payloadOf` text branch |
| `src/excel/links.ts` (modify) | `exportSelectionAsText`, the one-cell / empty / cap rules |
| `src/ppt/texts.ts` (create) | `insertText`, `refreshText`, `textSize` |
| `src/ppt/host.ts` (modify) | dispatch on `payload.kind === "text"` in `insertLink` and `refreshLink` |
| `src/ppt/change-source.ts` (modify) | candidates filtered by kind family |
| `src/pane/links-tab.ts` (modify) | button wiring, `exportRange(tab, kind)`, boot touch |
| `taskpane.html`, `src/help/copy.excel.ts` (modify) | the "Export as text" button and its sentence |
| `test/ppt.support.ts` (modify) | `seedText`, `pushText` |
| `test/links.text.integration.test.ts`, `test/links.touch.integration.test.ts`, `test/ppt.texts.integration.test.ts` (create) | the suites |
| docs (modify) | FEATURES.md, README.md (one sentence), Latvian guide, manual, CLAUDE.md Map, spec amendments |

---

### Task 1: Touch the workbook's links on Links-tab boot

**Files:**
- Create: `src/excel/link-touch.ts`
- Modify: `src/pane/links-tab.ts` (`boot(tab)`)
- Test: `test/links.touch.integration.test.ts`

**Interfaces:**
- Consumes: `readRegistry(context)` from `src/excel/link-anchors.ts` (returns `Registry`), `deriveLinkKeys(token)` from `src/link/crypto.ts` (returns `{ enc, auth }`), `RelayApi.touchLinks(items: TouchQuery[]): Promise<number>` from `src/link/relay.ts` (already shipped, see `src/link/relay.touch.test.ts`).
- Produces: `touchWorkbookLinks(relay: RelayApi): Promise<number>` (links touched; 0 without a request when the registry is empty).

- [ ] **Step 1: Write the failing test**

Create `test/links.touch.integration.test.ts` with the same boot as `test/links.table.integration.test.ts` (copy its imports, `memoryStore`, `beforeEach`/`afterEach`; seed `Model!B4:C5` with two rows and select it), then:

```ts
describe("touchWorkbookLinks", () => {
  it("touches every link the registry holds, with its own auth", async () => {
    const { touchWorkbookLinks } = await import("../src/excel/link-touch");
    await links.exportSelection(ws, relay);
    const touched = await touchWorkbookLinks(relay);
    expect(touched).toBe(1);
    expect(relay.touched).toEqual([expect.objectContaining({ id: expect.any(String) })]);
  });

  it("asks the relay nothing for a workbook without links", async () => {
    const { touchWorkbookLinks } = await import("../src/excel/link-touch");
    expect(await touchWorkbookLinks(relay)).toBe(0);
    expect(relay.touched).toEqual([]);
  });
});
```

`relay.touched` is a recorder to add to `test/fakerelay.ts`: an array every `touchLinks` call appends its items to (look at how the fake already records other calls; if it records none, add `readonly touched: TouchQuery[] = []` and push in `touchLinks`).

- [ ] **Step 2: Run it to make sure it fails**

Run: `npx vitest run test/links.touch.integration.test.ts`
Expected: FAIL, `Cannot find module '../src/excel/link-touch'`.

- [ ] **Step 3: Implement**

`src/excel/link-touch.ts`:

```ts
// Boot-time TTL refresh: tells the relay which links this workbook still holds
// so their 30 days run from today, not from the last push. Reads the registry
// only; never writes it, never toasts, never blocks the pane.
import { deriveLinkKeys } from "../link/crypto";
import type { RelayApi, TouchQuery } from "../link/relay";
import { readRegistry } from "./link-anchors";

export async function touchWorkbookLinks(relay: RelayApi): Promise<number> {
  const registry = await Excel.run((context) => readRegistry(context));
  if (registry.links.length === 0) return 0;
  const items: TouchQuery[] = [];
  for (const entry of registry.links) {
    const keys = await deriveLinkKeys(entry.token);
    items.push({ id: entry.id, auth: keys.auth });
  }
  return relay.touchLinks(items);
}
```

If `readRegistry` is not exported from `link-anchors.ts`, export it (it is read-only). In `src/pane/links-tab.ts` `boot(tab)`, after the existing `refresh(tab)` call:

```ts
  // Best effort: a failed touch changes nothing the user can see, and the next
  // boot tries again. Never a toast on boot.
  try {
    await touchWorkbookLinks(tab.deps.relay);
  } catch {
    return;
  }
```

with `import { touchWorkbookLinks } from "../excel/link-touch";`.

- [ ] **Step 4: Run the suite and the gate**

Run: `npx vitest run test/links.touch.integration.test.ts` then `npm run check`.
Expected: both PASS.

- [ ] **Step 5: Commit**

```bash
git add src/excel/link-touch.ts src/pane/links-tab.ts test/fakerelay.ts test/links.touch.integration.test.ts
git commit -m "Links tab touches the workbook's links on boot" -m "The relay's TTL now runs from the last pane boot, not the last push:
a model nobody exported for a month keeps repainting its decks.
Best effort, no toast."
```

---

### Task 2: The text payload in the model

**Files:**
- Modify: `src/link/model.ts` (lines 9, 43-92, 102, 359)
- Test: `src/link/model.test.ts` (extend; create beside `model.ts` if absent)

**Interfaces:**
- Produces: `LinkKind = "range" | "chart" | "table" | "text"`; `interface TextPayload { v: 1; kind: "text"; text: string; src: Source; pushedAt: string; hash: string }`; `Payload = PicturePayload | TablePayload | TextPayload`; `export const TEXT_MAX_CHARS = 500`; `export const TEXT_TOO_LONG = "Text links carry up to 500 characters. Export a longer cell as a picture."`; `sourceLabel(src, "text")` returns `` `${src.sheet}!${src.ref} text` ``.

- [ ] **Step 1: Write the failing tests**

```ts
import { decodePayload, encodePayload, payloadBytes, sourceLabel, TEXT_MAX_CHARS, type TextPayload } from "./model";

const text: TextPayload = {
  v: 1, kind: "text", text: "EUR 15.7m",
  src: { workbook: "Model_v4.xlsx", sheet: "P&L", ref: "C4", anchor: "PLSFIX_LINK_0123abcd" },
  pushedAt: "2026-08-31T09:00:00.000Z", hash: "0".repeat(64),
};

describe("text payload", () => {
  it("round-trips through the codec", () => {
    expect(decodePayload(encodePayload(text))).toEqual(text);
  });
  it("weighs its text", () => {
    expect(payloadBytes(text)).toBe(new TextEncoder().encode("EUR 15.7m").length);
  });
  it("is refused by the guard without a string text", () => {
    const bad = new TextEncoder().encode(JSON.stringify({ ...text, text: 42 }));
    expect(() => decodePayload(bad)).toThrow("not a link payload");
  });
  it("labels the cell and says text", () => {
    expect(sourceLabel(text.src, "text")).toBe("P&L!C4 text");
  });
  it("caps at 500 characters", () => {
    expect(TEXT_MAX_CHARS).toBe(500);
  });
});
```

- [ ] **Step 2: Run to verify failure**

Run: `npx vitest run src/link/model.test.ts`
Expected: FAIL on the type import and `sourceLabel` ("P&L!C4" without " text").

- [ ] **Step 3: Implement**

In `src/link/model.ts`: extend `LinkKind`; add after `TablePayload`:

```ts
// One cell's displayed text, nothing else: the slide decides the font. The
// hash is over the text, so an unchanged cell is an unchanged link.
export interface TextPayload {
  v: 1;
  kind: "text";
  text: string;
  src: Source;
  pushedAt: string;
  hash: string;
}

export type Payload = PicturePayload | TablePayload | TextPayload;

// A cell longer than this is a paragraph: a picture or a table says it better.
export const TEXT_MAX_CHARS = 500;
export const TEXT_TOO_LONG =
  "Text links carry up to 500 characters. Export a longer cell as a picture.";
```

`payloadBytes`: add `if (payload.kind === "text") return new TextEncoder().encode(payload.text).length;` before the existing branches. Find the guard `decodePayload` uses (`git grep -n "function isPayload" src`) and add the text case: `kind === "text"` requires `typeof value.text === "string"` plus the shared `src`/`pushedAt`/`hash` checks the other kinds make. `sourceLabel`: 

```ts
export function sourceLabel(src: Source, kind: LinkKind): string {
  if (kind === "chart") return `${src.sheet}: ${src.ref}`;
  const range = `${src.sheet}!${src.ref}`;
  if (kind === "table") return `${range} table`;
  return kind === "text" ? `${range} text` : range;
}
```

- [ ] **Step 4: Run tests + gate**

Run: `npx vitest run src/link/model.test.ts` then `npm run typecheck` (the union growth will point at every switch that must learn the new kind: fix only compile errors here, by treating `text` like `table` where a picture-only property is read; the real branches come in Tasks 3 and 4).
Expected: PASS, tsc clean.

- [ ] **Step 5: Commit**

```bash
git add src/link/model.ts src/link/model.test.ts
git commit -m "Text payload: one cell's displayed text" -m "Third payload kind beside picture and table. 500-character cap, label
'Sheet!A1 text' so a picture and a text of the same cell stay apart."
```

---

### Task 3: Excel exports one cell as text

**Files:**
- Modify: `src/excel/link-anchors.ts:35` (`ResolvedRange.kind`), `:53-56` (`Render`), `:337-350` (`renderSource`)
- Modify: `src/excel/link-record.ts:46-80` (`payloadOf`)
- Modify: `src/excel/links.ts:68-142` (`exportSelectionAsText`, `requireExportable`, `exportRange`)
- Test: `test/links.text.integration.test.ts`

**Interfaces:**
- Consumes: `TextPayload`, `TEXT_MAX_CHARS`, `TEXT_TOO_LONG` (Task 2); `sha256Hex` (already imported in `link-record.ts`).
- Produces: `exportSelectionAsText(ws: Workspace, relay: RelayApi): Promise<ExportResult>`; `Render` gains `{ kind: "text"; text: string }`; `ResolvedRange.kind: "range" | "table" | "text"`.

- [ ] **Step 1: Write the failing tests**

`test/links.text.integration.test.ts`, booted exactly like `test/links.table.integration.test.ts` (same imports, `memoryStore`, `beforeEach` installing the fake host with sheets `["Model", "Data"]`, `afterEach`), seeding `helpers.seed("Model!B4", [["EUR 15.7m", 1000]])` and selecting `Model!B4`:

```ts
function registryEntry(): { kind: string; label: string; token: string; id: string } {
  return JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links[0];
}

async function payloadOf(id: string, token: string): Promise<Payload> {
  const keys = await deriveLinkKeys(token);
  const result = await relay.getLink(id, keys.auth);
  if (result === "unchanged") throw new Error("no blob");
  return decodePayload(await open(keys.enc, id, result.blob));
}

describe("Export as text", () => {
  it("sends the displayed text of one cell under a text label", async () => {
    const { label } = await links.exportSelectionAsText(ws, relay);
    expect(label).toBe("Model!B4 text");
    const entry = registryEntry();
    expect(entry.kind).toBe("text");
    expect(helpers.names()).toContain(anchorName(entry.id));
    const payload = await payloadOf(entry.id, entry.token);
    expect(payload.kind).toBe("text");
    expect(payload.kind === "text" && payload.text).toBe("EUR 15.7m");
  });

  it("refuses two cells", async () => {
    helpers.select("Model!B4:C4");
    await expect(links.exportSelectionAsText(ws, relay)).rejects.toThrow("Select one cell");
    expect(helpers.setting(REGISTRY_SETTING)).toBeNull();
  });

  it("refuses an empty cell", async () => {
    helpers.select("Model!B9");
    await expect(links.exportSelectionAsText(ws, relay)).rejects.toThrow("empty");
  });

  it("refuses more than 500 characters before anchoring", async () => {
    helpers.seed("Model!B6", [["x".repeat(501)]]);
    helpers.select("Model!B6");
    await expect(links.exportSelectionAsText(ws, relay)).rejects.toThrow("500 characters");
    expect(helpers.names()).toEqual([]);
  });

  it("re-renders the text through the anchor on push", async () => {
    await links.exportSelectionAsText(ws, relay);
    const entry = registryEntry();
    helpers.seed("Model!B4", [["EUR 16.1m"]]);
    await links.pushLinks([entry.id], relay);
    const payload = await payloadOf(entry.id, entry.token);
    expect(payload.kind === "text" && payload.text).toBe("EUR 16.1m");
  });
});
```

If `helpers.names()` or `helpers.setting()` are named differently in `test/fakehost.ts`, use the fake's real accessors (read the top of `links.table.integration.test.ts` and `test/fakehost.ts` `FakeHelpers`). If `pushLinks` takes a different shape (see its signature at `src/excel/links.ts:266`), call it the way `links.integration.test.ts` does.

- [ ] **Step 2: Run to verify failure**

Run: `npx vitest run test/links.text.integration.test.ts`
Expected: FAIL, `exportSelectionAsText is not a function`.

- [ ] **Step 3: Implement**

`src/excel/link-anchors.ts`: `kind: "range" | "table" | "text";` on `ResolvedRange`; `Render` gains `| { kind: "text"; text: string }`; in `renderSource`, before the picture fallthrough:

```ts
  if (resolved.kind === "text") {
    resolved.range.load("text");
    await context.sync();
    return { kind: "text", text: resolved.range.text[0]?.[0] ?? "" };
  }
```

`src/excel/link-record.ts` `payloadOf`, before the picture return:

```ts
  if (render.kind === "text") {
    return {
      v: 1,
      kind: "text",
      text: render.text,
      src,
      pushedAt,
      hash: await sha256Hex(render.text),
    };
  }
```

`src/excel/links.ts`:

```ts
// One cell's displayed text as a text box on the slide: same anchor, same
// registry entry, same inbox note as a picture of it.
export async function exportSelectionAsText(
  ws: Workspace,
  relay: RelayApi,
): Promise<ExportResult> {
  return exportRange(ws, relay, "text");
}
```

`exportRange`'s `kind` parameter becomes `"range" | "table" | "text"` (the `exclusive` stage: `kind === "range" ? "export" : \`export ${kind}\``). Load `text` alongside the other properties only for text (`range.load(kind === "text" ? "address,cellCount,rowCount,columnCount,worksheet/name,text" : "address,cellCount,rowCount,columnCount,worksheet/name")`), then extend `requireExportable(range, kind)`:

```ts
  if (kind === "text") {
    if (range.cellCount !== 1) {
      throw new Error("Select one cell for a text link (merged cells: export as a picture).");
    }
    const text = range.text[0]?.[0] ?? "";
    if (text.trim() === "") throw new Error("The cell is empty.");
    if (text.length > TEXT_MAX_CHARS) throw new Error(TEXT_TOO_LONG);
  }
```

Keep `requireExportable` under 50 lines: if it grows past it, move the text rules into `function requireTextCell(range: Excel.Range): void` beside it.

- [ ] **Step 4: Run tests + gate**

Run: `npx vitest run test/links.text.integration.test.ts test/links.table.integration.test.ts test/links.integration.test.ts` then `npm run check`.
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/excel/link-anchors.ts src/excel/link-record.ts src/excel/links.ts test/links.text.integration.test.ts
git commit -m "Export one cell as a text link" -m "Range.text of a single cell travels as the text payload; two cells, an
empty cell or more than 500 characters are refused before the anchor is
made. Push re-renders through the same anchor."
```

---

### Task 4: PowerPoint inserts and refreshes a text box

**Files:**
- Create: `src/ppt/texts.ts`
- Modify: `src/ppt/host.ts:184-233` (`insertLink`), `:260-300` (`refreshLink`), imports at the top
- Modify: `test/ppt.support.ts` (add `seedText`, `pushText` beside `seedTable`/`pushTable`)
- Test: `test/ppt.texts.integration.test.ts`

**Interfaces:**
- Consumes: `selectedSlideId(context, stage)`, `placeOnSlide(context, slideId, size)` -> `{ box, overlapping }` and `CONTENT_WIDTH` from `src/ppt/placement.ts`; `shapeAt(context, found)` from `src/ppt/shapes.ts`; `encodeTag`, `TAG_LINK`, `TAG_KEY`, `sourceLabel` from `src/link/model.ts`; `Size` from `src/layout.ts`; `InsertResult`, `FoundLink`, `tagFor` from `host.ts` (export `tagFor` if it is private, or pass the tag in as `insertTable` does).
- Produces: `insertText(stage: string, item: InboxItem, payload: TextPayload, tag: LinkTag): Promise<InsertResult>`; `refreshText(found: FoundLink, payload: TextPayload, tag: LinkTag): Promise<void>`; `textSize(payload: TextPayload): Size`; `seedText(text: string): Promise<InboxItem>`; `pushText(item: InboxItem, text: string): Promise<void>`.

- [ ] **Step 1: Add the seed helpers**

In `test/ppt.support.ts`, beside `seedTable`:

```ts
// A text export waiting in the inbox: the label is the one sourceLabel gives.
export async function seedText(text: string): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload: Payload = {
    v: 1, kind: "text", text, src,
    pushedAt: new Date().toISOString(), hash: "0".repeat(64),
  };
  await publish(id, token, payload);
  return { id, token, kind: "text", label: `${src.sheet}!${src.ref} text`, src: payload.src, createdAt: payload.pushedAt };
}

export async function pushText(item: InboxItem, text: string): Promise<void> {
  await publish(item.id, item.token, {
    v: 1, kind: "text", text, src,
    pushedAt: new Date().toISOString(), hash: "1".repeat(64),
  });
}
```

- [ ] **Step 2: Write the failing tests**

`test/ppt.texts.integration.test.ts`, booted with `bootPpt()` like `test/ppt.table.integration.test.ts` (copy its `beforeEach`, workspace creation with `memoryStore()`, and its list -> update idiom):

```ts
describe("text links", () => {
  it("inserts a text box with the text, auto-sized, tagged, in free space", async () => {
    const item = await seedText("EUR 15.7m");
    const placed = await links.insertFromInbox(item, ws, relay);
    const shape = helpers.shape(placed.shapeId);
    expect(shape.type).toBe("TextBox");
    expect(shape.text).toBe("EUR 15.7m");
    expect(shape.autoSize).toBe("AutoSizeShapeToFitText");
    expect(shape.wordWrap).toBe(false);
    expect(shape.tags.get(TAG_LINK)).toBeDefined();
    expect(shape.tags.get(TAG_KEY)).toBe(item.token);
    expect(placed.overlapping).toBe(false);
  });

  it("refreshes the text in place and leaves the box where it was", async () => {
    const item = await seedText("EUR 15.7m");
    const placed = await links.insertFromInbox(item, ws, relay);
    const shape = helpers.shape(placed.shapeId);
    helpers.moveShape(placed.shapeId, { left: 100, top: 200, width: 150 });
    await pushText(item, "EUR 16.1m");
    await updateAll();
    expect(shape.text).toBe("EUR 16.1m");
    expect([shape.left, shape.top, shape.width]).toEqual([100, 200, 150]);
    expect(decodeTag(shape.tags.get(TAG_LINK))?.rev).toBe(2);
  });

  it("costs no more syncs than a picture insert", async () => {
    const picture = await seedLink(PNG_1x1);
    const before = helpers.syncCount();
    await links.insertFromInbox(picture, ws, relay);
    const pictureSyncs = helpers.syncCount() - before;
    const item = await seedText("EUR 15.7m");
    const start = helpers.syncCount();
    await links.insertFromInbox(item, ws, relay);
    expect(helpers.syncCount() - start).toBeLessThanOrEqual(pictureSyncs);
  });

  it("reverts to the previous text", async () => {
    const item = await seedText("EUR 15.7m");
    const placed = await links.insertFromInbox(item, ws, relay);
    await pushText(item, "EUR 16.1m");
    await updateAll();
    await revertSelected([placed.shapeId]);
    expect(helpers.shape(placed.shapeId).text).toBe("EUR 15.7m");
  });

  it("breaks the link and leaves a plain text box", async () => {
    const item = await seedText("EUR 15.7m");
    const placed = await links.insertFromInbox(item, ws, relay);
    await breakSelected([placed.shapeId]);
    const shape = helpers.shape(placed.shapeId);
    expect(shape.text).toBe("EUR 15.7m");
    expect(shape.tags.get(TAG_LINK)).toBeUndefined();
  });
});
```

`helpers.shape`, `helpers.moveShape`, `helpers.syncCount`, `updateAll`, `revertSelected`, `breakSelected`, `PNG_1x1` are the fake's and the sibling suites' own names: take them from `test/fakeppt/index.ts` (`FakePptHelpers`), `test/ppt.update.integration.test.ts`, `test/ppt.revert.integration.test.ts` and `test/ppt.table.integration.test.ts`, and adapt the calls to their real signatures; do not invent new helpers when one exists. The fake shape must expose `text`, `autoSize`, `wordWrap` (see `test/fakeppt/model.ts` and `format.ts`; add a field only if it is missing).

- [ ] **Step 3: Run to verify failure**

Run: `npx vitest run test/ppt.texts.integration.test.ts`
Expected: FAIL: the insert lands a picture path or throws on `payload.png` for a text payload.

- [ ] **Step 4: Implement `src/ppt/texts.ts`**

```ts
// A text link on a slide: one tagged text box of its own, auto-sized to what
// Excel displayed, placed in free space like every other kind. Owns the insert
// and the in-place refresh; a refresh writes the text and the tag, never the
// geometry, font or colour the user chose.
import type { Size } from "../layout";
import {
  encodeTag,
  sourceLabel,
  TAG_KEY,
  TAG_LINK,
  type InboxItem,
  type LinkTag,
  type TextPayload,
} from "../link/model";
import type { FoundLink, InsertResult } from "./host";
import { CONTENT_WIDTH, placeOnSlide, selectedSlideId } from "./placement";
import { shapeAt } from "./shapes";

// PowerPoint's default text-box font is 18 pt; 0.55 em per character is the
// width of a proportional font's average glyph. The auto-size corrects it.
const FONT_PT = 18;
const EM_PER_CHAR = 0.55;
const LINE_HEIGHT = 28;
const MIN_WIDTH = 60;

export function textSize(payload: TextPayload): Size {
  const width = Math.ceil(payload.text.length * FONT_PT * EM_PER_CHAR) + 14;
  return {
    width: Math.min(CONTENT_WIDTH, Math.max(MIN_WIDTH, width)),
    height: LINE_HEIGHT,
  };
}

export async function insertText(
  stage: string,
  item: InboxItem,
  payload: TextPayload,
  tag: LinkTag,
): Promise<InsertResult> {
  return PowerPoint.run(async (context) => {
    const slideId = await selectedSlideId(context, stage);
    const placed = await placeOnSlide(context, slideId, textSize(payload));
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    const shape = shapes.addTextBox(payload.text, placed.box);
    shape.name = `pls,fix text ${item.label}`;
    shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
    shape.textFrame.wordWrap = false;
    shape.tags.add(TAG_LINK, encodeTag(tag));
    shape.tags.add(TAG_KEY, item.token);
    shape.load("id");
    await context.sync();
    return { slideId, shapeId: shape.id, overlapping: placed.overlapping };
  });
}

// The text the source shows now, written into the box where it sits.
export async function refreshText(
  found: FoundLink,
  payload: TextPayload,
  tag: LinkTag,
): Promise<void> {
  const stage = `refresh ${sourceLabel(found.tag.src, found.tag.kind)}`;
  await PowerPoint.run(async (context) => {
    const shape = shapeAt(context, found);
    shape.textFrame.textRange.text = payload.text;
    shape.tags.add(TAG_LINK, encodeTag(tag));
    await context.sync();
  }).catch((error: unknown) => {
    throw new Error(`${stage}: ${error instanceof Error ? error.message : String(error)}`);
  });
}
```

If the fake host does not know `PowerPoint.ShapeAutoSize.autoSizeShapeToFitText`, use the string `"AutoSizeShapeToFitText"` the way the codebase writes other enums (check `git grep -n "GeometricShapeType" src/ppt`). `src/ppt/host.ts`: in `insertLink`, before the table branch:

```ts
  if (payload.kind === "text") {
    return insertText(stage, item, payload, tag);
  }
```

and in `refreshLink`, before the table branch:

```ts
  if (payload.kind === "text") {
    await refreshText(found, payload, tagFor(found.tag, payload, rev));
    return;
  }
```

`refreshLinks` (the batch) already returns `false` for a batch that is not all pictures, so Update all takes the per-row path for text rows. Update the header comment of `host.ts` ("texts.ts the text box").

- [ ] **Step 5: Run tests + gate**

Run: `npx vitest run test/ppt.texts.integration.test.ts test/ppt.update.integration.test.ts test/ppt.revert.integration.test.ts test/ppt.perf.integration.test.ts` then `npm run check`.
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/ppt/texts.ts src/ppt/host.ts test/ppt.support.ts test/ppt.texts.integration.test.ts test/fakeppt
git commit -m "Text links land as tagged text boxes" -m "addTextBox (PowerPointApi 1.4), auto-sized, placed in free space; a
refresh writes the text and the tag only. Revert, break and Update all
work through the per-row path like tables."
```

---

### Task 5: Change source stays inside the kind family

**Files:**
- Modify: `src/ppt/change-source.ts:53-58` (`candidatesFor`)
- Test: `test/ppt.change-source.integration.test.ts` (extend) or `src/ppt/change-source.test.ts` if the pure part is tested there (check with `ls src/ppt/*.test.ts`)

**Interfaces:**
- Produces: `candidatesFor(row: LinkRow, inbox: InboxItem[]): InboxItem[]` now drops items across the text boundary: a text row sees only `kind === "text"` items; a picture or table row never sees a text item.

- [ ] **Step 1: Write the failing test**

```ts
it("offers a text link only text exports, and hides text from the rest", () => {
  const textRow = rowWithKind("text");
  const pictureRow = rowWithKind("range");
  const inbox = [inboxItem("text", "P&L!C4 text"), inboxItem("range", "P&L!C4"), inboxItem("table", "P&L!C4:D9 table")];
  expect(candidatesFor(textRow, inbox).map((i) => i.kind)).toEqual(["text"]);
  expect(candidatesFor(pictureRow, inbox).map((i) => i.kind)).toEqual(["range", "table"]);
});
```

Build `rowWithKind` / `inboxItem` from the fixtures the existing change-source tests already use (same file), changing only `kind` and `label`.

- [ ] **Step 2: Run to verify failure**

Run: `npx vitest run test/ppt.change-source.integration.test.ts` (or the pure test file)
Expected: FAIL: the text row receives all three.

- [ ] **Step 3: Implement**

```ts
// A text box cannot become a picture in place, nor the other way round: the
// candidates stay on the row's side of that line. Picture and table still mix,
// as before (kindWarning says so when they do).
function sameFamily(row: LinkRow, item: InboxItem): boolean {
  return (row.found.tag.kind === "text") === (item.kind === "text");
}

export function candidatesFor(row: LinkRow, inbox: InboxItem[]): InboxItem[] {
  return inbox
    .map((item, index) => ({ item, rank: rank(row, item), index }))
    .filter((entry) => sameFamily(row, entry.item))
    .sort((left, right) => left.rank - right.rank || left.index - right.index)
    .map((entry) => entry.item);
}
```

- [ ] **Step 4: Run tests + gate**

Run: the test file, then `npm run check`. Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/ppt/change-source.ts test/ppt.change-source.integration.test.ts
git commit -m "Change source keeps text links among text exports" -m "A tagged text box cannot take a picture's payload in place, so the
picker no longer offers one across that line."
```

---

### Task 6: The pane: button, sentence, wiring

**Files:**
- Modify: `taskpane.html:1240-1264` (the Export action list), `src/help/copy.excel.ts:233-244` (`links["export-heading"].buttons`), `src/pane/links-tab.ts` (`wire(...)` calls and `exportRange(tab, ...)`)
- Test: `src/help/copy.test.ts` (existing; it parses the HTML and fails on a button without a sentence), `npm run ux:check`

- [ ] **Step 1: Run the copy test to see it pass today**

Run: `npx vitest run src/help/copy.test.ts`. Expected: PASS (baseline).

- [ ] **Step 2: Add the button**

After the `export-table` button in `taskpane.html`:

```html
            <button id="export-text" type="button">
              <span class="action-icon links">T</span
              ><span
                ><strong>Export as text</strong
                ><small>One cell's text as a box that refreshes in place</small></span
              >
            </button>
```

- [ ] **Step 3: Run the copy test to see it fail**

Run: `npx vitest run src/help/copy.test.ts`. Expected: FAIL naming `export-text`.

- [ ] **Step 4: Add the sentence and the wiring**

`src/help/copy.excel.ts`, in `links["export-heading"].buttons`:

```ts
      "export-text":
        "Sends one cell's displayed text as a text box that keeps its place, size and font when it refreshes. Up to 500 characters; the slide decides the look.",
```

`src/pane/links-tab.ts`: change `exportRange(tab, asTable: boolean)` to take the kind:

```ts
type ExportKind = "range" | "table" | "text";

async function exportRange(tab: Tab, kind: ExportKind): Promise<string> {
  const send =
    kind === "table"
      ? exportSelectionAsTable
      : kind === "text"
        ? exportSelectionAsText
        : exportSelection;
  const result = await send(requireWorkspace(tab), tab.deps.relay);
  await refresh(tab);
  return `Sent to PowerPoint: ${result.label}`;
}
```

and the three wire calls become `exportRange(tab, "range")`, `exportRange(tab, "table")`, plus `wire(tab, "export-text", () => exportRange(tab, "text"));`. Import `exportSelectionAsText` from `../excel/links` (or through the barrel the file already uses).

- [ ] **Step 5: Gates**

Run: `npx vitest run src/help/copy.test.ts`, `npm run check`, then with `npm run dev` running in another terminal `npm run ux:check`.
Expected: all PASS; ux:check reports 0 defects at 320/360/420/500 (a fifth action button must not clip or scroll sideways; if it does, the fix belongs in `src/styles/excel.css` for the action list, nothing else).

- [ ] **Step 6: Commit**

```bash
git add taskpane.html src/help/copy.excel.ts src/pane/links-tab.ts
git commit -m "Links tab: Export as text" -m "Fourth export button with its help sentence; the tab's export entry
takes the kind instead of a table flag."
```

---

### Task 7: Docs, Map and spec amendments

**Files:**
- Modify: `docs/FEATURES.md:149`, `README.md` (the "Export as table" sentence in "Linked objects in PowerPoint", around line 147), `docs/lietotaja-rokasgramata-saites.md` (section "## Eksports no Excel", line 10), `manual/src/content/links.rs` (`fn export()` section, line 67), `CLAUDE.md` Map, `docs/superpowers/specs/2026-08-31-text-links-design.md`

- [ ] **Step 1: FEATURES.md**

Row 149 becomes:

```
| Export text (cell -> text box) | both | Live text links | shipped v2.5 (one cell, whole-box; runs inside a sentence deferred) | P2 (M3) |
```

- [ ] **Step 2: README.md (hand-curated: this one sentence only)**

After the sentence ending "(up to 60 rows and 20 columns, PowerPoint 2021 or Microsoft 365)." add:

```
*Export as text* sends one cell's displayed text as a text box that keeps its place, size and font when it refreshes (up to 500 characters).
```

- [ ] **Step 3: Latvian guide**

Under "## Eksports no Excel", after the bullet that describes the table export, add:

```
- *Export as text* nosūta vienas šūnas attēloto tekstu (skaitli vai frāzi) kā teksta lauku, ko PowerPoint atjauno tajā pašā vietā; lauka izmērs un fonts paliek tādi, kādus tos atstājāt (līdz 500 rakstzīmēm).
```

- [ ] **Step 4: Manual**

In `fn export()` of `manual/src/content/links.rs`, after the paragraph about the table export, add a block:

```rust
            Block::Para(
                r#"Ar "Export as text" vienas šūnas attēlotais teksts nonāk slaidā kā atsevišķs teksta lauks. Atjaunināšana maina tikai tekstu; lauka vieta, izmērs un fonts paliek tādi, kādus tos atstājāt (līdz 500 rakstzīmēm)."#,
            ),
```

Run `cargo test --manifest-path manual/Cargo.toml` if the manual crate has tests, and `npm run manual` only if `manual/out` is tracked (check `git ls-files manual/out`; if it is not tracked, do not regenerate).

- [ ] **Step 5: CLAUDE.md Map**

- `src/excel/` line: after `link-lock.ts (...)` add `+ link-touch.ts (boot-time TTL refresh: registry -> touchLinks)`; in the `links.ts` clause mention `exportSelectionAsText`.
- `src/ppt/` line: add `texts.ts (a text link as one tagged text box: addTextBox on PowerPointApi 1.4, auto-sized, placed in free space; refresh writes the text and the tag only)`.
- `src/link/` line: the payload union is `a picture, a table ..., or a text (one cell's displayed text, TEXT_MAX_CHARS 500)`.
- Data flow line: `or src/excel/link-text` is not a file; say `or the range's text for a text link`.
- "Where bugs live": add `a text link's box moved, resized or restyled on update -> src/ppt/texts.ts (a refresh may write the text and the tag only)`.
- Keep the Map at most 40 lines.

- [ ] **Step 6: Spec amendments (so the spec matches what shipped)**

In `docs/superpowers/specs/2026-08-31-text-links-design.md`: (a) Excel section: replace the merged-area rule with "v1: exactly one cell (`cellCount === 1`); a merged area is refused with 'Select one cell for a text link (merged cells: export as a picture)'; the ExcelApi 1.13 merged-area rule is a follow-up"; (b) the label reads `Sheet!A1 text`, not `Sheet!A1`, so a picture and a text of the same cell stay apart in the lists; (c) add the boot-time `touchWorkbookLinks` (Task 1) under a short "Shipped with it" heading.

- [ ] **Step 7: Gate + commit**

Run: `npm run check`. Expected: PASS.

```bash
git add docs/FEATURES.md README.md docs/lietotaja-rokasgramata-saites.md manual/src/content/links.rs CLAUDE.md docs/superpowers/specs/2026-08-31-text-links-design.md
git commit -m "Text links: docs, Map and spec as shipped" -m "FEATURES row shipped, one README sentence, the Latvian guide and manual
paragraph, Map lines for link-touch.ts and texts.ts, spec amended to the
one-cell rule and the 'text' label."
```

---

## Self-review (done while writing)

- Spec coverage: decisions 1-6 map to Tasks 2-6; data model to Task 2; Excel to Task 3; PowerPoint to Task 4; change source to Task 5; pane to Task 6; docs to Task 7; the merged-area rule is consciously narrowed and recorded in Task 7 step 6; the web-rig verification is the controller's, after merge, when a signed-in Chrome is available.
- Types: `TextPayload`, `TEXT_MAX_CHARS`, `TEXT_TOO_LONG`, `exportSelectionAsText`, `insertText`, `refreshText`, `textSize`, `seedText`, `pushText`, `touchWorkbookLinks`, `sameFamily`, `ExportKind` are named identically wherever they appear.
- Placeholders: helper names taken from sibling suites are called out as such with the file to read; no "TBD".
