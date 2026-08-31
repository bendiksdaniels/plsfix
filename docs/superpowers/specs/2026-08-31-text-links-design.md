# Text links: one cell to a slide text box (design, 31.08.2026)

## Why

The third link kind. A number in a slide title, a KPI callout, a footnote
figure: today those are retyped after every model change. Picture and table
links cover blocks; a single cell exported as a picture is useless inside text.
FEATURES.md section 8 lists "Export text" as the open P2 (M3) item, and it is
the UpSlide feature analysts use most after tables.

## Decisions (final)

1. Kind `text`: exactly one cell. The payload carries the displayed text
   (`Range.text[0][0]`, what the exporter's Excel shows, number format and
   separators included) and nothing else. No font, no fill, no size: the slide
   decides how the text looks.
2. PowerPoint: a tagged text box of its own (`shapes.addTextBox`,
   PowerPointApi 1.4, so every host that runs the pane can do it), auto-sized
   to its text, placed in free space like the other kinds. A refresh writes
   `textFrame.textRange.text` only: position, size, font, colour and alignment
   stay the user's. Whole-box semantics: the box holds exactly the linked text.
3. Out of scope: a run inside an existing sentence. PowerPoint JS has no
   anchored run; character offsets shift with every edit, so a refresh would
   hit the wrong characters. The tagged box beside the sentence is the robust
   version; a run link needs its own spike on `TextRange` anchoring first.
4. Auto-push, highlight, revert, change source, break and Update all stay as
   they are: they act on tags, anchors and payload revisions. Text rows repaint
   through the per-row path like tables (a text repaint is a few bytes; the
   byte-budgeted picture batch gains nothing from it).
5. No house number style is applied at export. The displayed text is the
   truth; the brand language shapes what the pane itself writes, not the link.
   The help card says so.
6. Change source: a text link may only re-point to another text export (a
   picture cannot become a text box in place). Candidates are filtered by kind.

## Data model (`src/link/model.ts`)

- `LinkKind` gains `"text"`.
- `TextPayload { v: 1; kind: "text"; text: string; src: Source; pushedAt:
  string; hash: string }`; `Payload` gains it; `payloadBytes` counts the text.
- `TEXT_MAX_CHARS` = 500. A longer cell is a paragraph: export it as a picture
  or a table. Refused before anything is anchored, like `TABLE_TOO_BIG`.
- Tag, registry and inbox codecs need no change beyond the kind.

## Excel (`src/excel/links.ts` + new `src/excel/link-text.ts`)

- `exportSelectionAsText(ws, relay)` -> `exportRange(ws, relay, "text")`.
- `requireExportable` for text, v1 as shipped: exactly one cell
  (`cellCount === 1`). A merged area selects as one range of several cells and
  is refused with "Select one cell for a text link (merged cells: export as a
  picture)."; the ExcelApi 1.13 merged-area rule
  (`getMergedAreasOrNullObject`) is a follow-up. The other two errors are "The
  cell is empty." and "Text links carry up to 500 characters. Export a longer
  cell as a picture.".
- `renderText(range)`: `range.load("text")`, the value is `text[0][0]`.
- Registry label and Links list: `sourceLabel(src, "text")` reads as
  `Sheet!A1 text`, not `Sheet!A1`, so a picture and a text of the same cell
  stay apart in both lists; the kind column says "text".
- `link-watch.ts` (auto-push) and `link-highlight.ts` need nothing new: the
  anchor is the hidden name, the tint covers the one cell.

## PowerPoint (new `src/ppt/texts.ts`, `host.ts` dispatch, `links.ts`)

- `insertText(stage, item, payload, tag)`: `selectedSlideId` ->
  `placeOnSlide(context, slideId, textSize(payload))` -> `shapes.addTextBox(
  payload.text, box)`; `shape.name = "pls,fix text <label>"`;
  `textFrame.autoSizeSetting = "AutoSizeShapeToFitText"`, `wordWrap = false`;
  both tags; `load("id")`; one sync. `textSize` estimates width from the
  character count (0.55 em at 18 pt, at least 60 pt, at most `CONTENT_WIDTH`)
  and one line of 28 pt for the height; the auto-size corrects it.
- `refreshText(found, payload, tag)`: the shape `found` names (through its
  `groupPath` when grouped) -> `textFrame.textRange.text = payload.text` and
  the `TAG_LINK` rewrite with the new rev; one sync; geometry untouched.
- `host.insertLink` / `host.refreshLink`: `payload.kind === "text"` branches
  before the picture path. `scanLinks` is unchanged (tags). `isPicture`
  already excludes non-picture payloads, so Update all sends text rows through
  `refreshLink`.
- `change-source.ts`: candidates filtered to the same kind as the ticked link.
- Revert: the same repaint with the rev-1 text.

## Pane

- Excel Links tab: button "Export as text" (`data-action="export-text"`)
  beside "Export selection" and "Export as table"; its "?" sentence in
  `src/help/copy.excel.ts` (the copy test enforces it); the list shows the kind.
- PowerPoint pane: inbox and link rows show the kind "text"; Insert dispatches
  on the payload; toast "Linked text inserted on slide N".

## Fake hosts and tests

- `test/fakeppt`: `addTextBox` and `autoSizeSetting` exist; record
  `textRange.text` writes; strict load rules on `textFrame`.
- `test/fakehost.ts` (Excel): `range.text` and `getMergedAreasOrNullObject`
  if missing.
- Tests: `src/link/model.test.ts` (codec round trip, cap); Excel integration
  (one cell ok, two cells refused, empty refused, merged area ok, 501 chars
  refused); `test/ppt.texts.integration.test.ts` (insert as a text box in
  free space; refresh writes only the text and the tag; revert; change source
  only to a text export; break leaves the box); `ppt.perf` budget: insert =
  placement syncs + 1, refresh = 1 sync; `src/help/copy.test.ts` green.

## Docs

- `docs/FEATURES.md` section 8: "Export text" -> shipped v2.5.
- `README.md` (hand-curated): one sentence under "Linked objects", Excel side.
- `docs/lietotaja-rokasgramata-saites.md` and the manual: a short "Teksta
  saite" paragraph.
- CLAUDE.md Map: `link-text.ts`, `texts.ts`, the data-flow line, a symptom line
  ("a text link's box moved or restyled on update -> `src/ppt/texts.ts`").

## Shipped with it

- Boot-time TTL refresh (`src/excel/link-touch.ts`): when the Links tab boots,
  `touchWorkbookLinks(relay)` names every link the registry still holds to the
  already shipped `POST /api/links/touch`, each with its own auth key, so the
  relay's 30 days run from the last pane boot rather than the last push. Best
  effort: a failure changes nothing the user can see and never toasts.
- `src/excel/links.ts` and `src/excel/link-anchors.ts` were both at the
  400-line cap, so the selection exports moved to `src/excel/link-export.ts`
  and every render to `src/excel/link-render.ts`. Pure moves; `links.ts`
  re-exports the export flows, so no import path outside `src/excel/` changed.

## Size and risks

- About 600-900 lines including tests; one worktree agent (opus: adapter work
  on both hosts), controller review, then the web rig.
- Risks: `Range.text` on a merged area (prove in the fake, then on the web);
  `autoSizeSetting` on PowerPoint for the web (a text box costs about 1 s
  there, one box is fine); `addTextBox` options are points like the rest.

## Verification

- Fake-host round trip in every suite, `npm run check`, `npm run ux:check`.
- Web rig (`scripts/rig`): export `Bridge!C4` as text, insert, change the
  cell, Push all, Update all: the box text changes and the box stays where it
  was; Revert brings the old text back; Break leaves a plain text box.
