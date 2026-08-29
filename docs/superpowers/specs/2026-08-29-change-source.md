# Change source (v2.1, backlog D7) - design

Status: design 2026-08-29. Scope: ROADMAP Milestone 3, "Change source and resolve workbook
versions". Goal: re-point a tracked picture at a different export - typically the same table
exported again from a newer workbook (`Model_v4.xlsx` -> `Model_v5.xlsx`) - keeping its slide,
position and size.

## 1. Data touched

| Where | What changes |
|---|---|
| `SMT_LINK` / `SMT_KEY` tags | rewritten whole: new `id`, `kind`, `rev`, `src`, `pushedAt`, and the new link's token |
| The picture | repainted with the new link's latest render |
| Geometry, slide, z-order, group | untouched; `height` only on an aspect change, by the existing refresh rule |
| The relay | the chosen inbox item is deleted, as an Insert would; the old link is left alone - still pushable from Excel, still reachable by another deck |
| Excel (registry, anchors, workbook) | nothing |

## 2. Candidates

`candidatesFor(row, inbox)` is pure and returns **every** inbox item, ordered so the likely one
is the default: exact `label` match first, then same `src.anchor` or same sheet + `src.ref`,
then the rest ("any inbox item" fallback). Comparison folds case, the way `sourceChanged` does;
order is stable inside a tier. The chooser shows `label - workbook - age`, so two exports of one
table are told apart.

## 3. Flow

1. Exactly one row ticked -> "Change source" is live (`requireOneRow` refuses 0 or 2+).
2. It opens the inline chooser in the Links view: a `<select>` of the candidates plus Confirm /
   Cancel. Nothing is remembered while it is open - Confirm reads the tick and the inbox again,
   so a rescan underneath the picker cannot re-point another picture.
3. Confirm -> `changeSource(row, item, ws, relay, host)`: `deriveLinkKeys(item.token)` ->
   `relay.getLink(item.id, auth)` (no `knownRev`: the deck has never held this link) -> `open`
   + `decodePayload` -> new `LinkTag` -> `host.retagLink` -> repaint through the shared
   `applyBatch` with the retargeted `FoundLink` -> `relay.deleteInbox`.
4. The pane re-reads the inbox and rescans; the row now reports `current` against the **new**
   link, because its tag holds the new id and the rev the relay just answered with.

Retag before repaint, not after: below PowerPointApi 1.8 the repaint reinserts the picture as a
new shape, so a retag afterwards would address a shape that is gone. The reinsertion carries
the new tags itself (it is handed the retargeted link), and on 1.8 `fill.setImage` never writes
`SMT_KEY` at all - `retagLink` is the only thing that does. A repaint that fails puts the old tags
back, so a failure never leaves a shape claiming to be current while showing the old picture.

## 4. Edge cases

- **Two or more rows ticked, or none**: refused before the chooser opens.
- **Empty inbox / not paired**: refused with "Nothing waiting in the Inbox ...".
- **Candidate of the other kind** (a chart for a range link, or the reverse): allowed, with a
  warning staged in the toast details. The tag's `kind` follows the new link.
- **The link the row already holds**: harmless, it repaints and re-tags itself. **A grouped
  picture below 1.8**: the existing refusal in `refreshLink` applies unchanged.
- **Relay 404 / wrong key / network**: the typed `RelayError` reaches the toast, deck untouched.

## 5. API and tests

```ts
// src/ppt/host.ts - Office.js
retagLink(found: FoundLink, tag: LinkTag, token: string): Promise<void>
// src/ppt/change-source.ts - no Office.js, host and relay injectable
candidatesFor(row: LinkRow, inbox: InboxItem[]): InboxItem[]   // pure, ordered
requireOneRow(rows: LinkRow[]): LinkRow                        // pure, refuses 0 or 2+
requireCandidates(items: InboxItem[]): InboxItem[]             // pure, refuses empty
pickCandidate(items: InboxItem[], id: string): InboxItem       // pure, refuses unknown
kindWarning(row: LinkRow, item: InboxItem): string | undefined // pure
changeSource(row, item, ws, relay, host?): Promise<string>     // "Source changed: A -> B"
// src/ppt/views.ts - DOM
renderCandidates(select: HTMLSelectElement, items: InboxItem[]): void
```

`PptHost` gains `retagLink`; every error carries its stage and link (`change source
Model!B4:F12: ...`). Pane wiring lives in `src/ppt/chooser.ts`, so `main.ts` gains one install
call rather than a fourth screen's worth of handlers.

`test/ppt.change-source.integration.test.ts` (fake deck + fake relay): a seeded link and two inbox
items, one the same label from `Model_v5.xlsx` - `candidatesFor` puts it first; `changeSource`
repaints with the new picture, both tags carry the new id, token and src, the box is unchanged, the
inbox item is gone, and a later `listLinks` reports the row `current` against the new link; refusal
with two ticked rows and with an empty inbox. Plus `renderCandidates` in `src/ppt/views.test.ts`.
