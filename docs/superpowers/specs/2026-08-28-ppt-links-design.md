# Excel -> PowerPoint tracked links (v2.0) - design

Status: approved 2026-08-28 (Daniel). Transport decision: suite Rust host as an end-to-end
encrypted relay. Scope: ROADMAP Milestone 2 in full + the core of Milestone 3.

## 1. Goal

A range or chart exported from Excel into PowerPoint carries a tracker on both ends, so it
can be refreshed later, in place, with position and size preserved, even after the object was
moved between slides / resized / copied, and even after the source range or chart moved
inside the workbook (rows inserted, sheet renamed, chart moved).

Non-goals for v2.0 (v2.1 candidates): native PowerPoint tables, change-source / workbook
version resolution, highlight linked cells in Excel, auto-push on edit, revert to previous
render, shapes inside groups, Entra SSO pairing.

## 2. Verified platform facts (learn.microsoft.com, 2026-08-28)

| Need | API | Requirement set |
|---|---|---|
| Durable identity on a shape | `Shape.tags` (TagCollection add/getItemOrNullObject/delete; keys uppercase, string values) | PowerPointApi 1.3 |
| Create the object without touching selection | `ShapeCollection.addGeometricShape(Rectangle,{left,top,width,height})` | 1.4 |
| Refresh in place | `Shape.fill.setImage(base64)` (fill type -> PictureAndTexture) | 1.8 (Win 2504, Mac 16.96; NOT on VL/LTSC) |
| Fallback insert below 1.8 | `Office.context.document.setSelectedDataAsync(png,{coercionType:Image,imageLeft,imageTop,imageWidth,imageHeight})` then last shape of the slide | Common API |
| Geometry, name (writable), delete, parent slide | `Shape.left/top/width/height/name`, `delete()`, `getParentSlideOrNullObject()` | 1.4 / 1.3 / 1.5 |
| Z-order | `Shape.zOrderPosition`, `setZOrder()` | 1.8 |
| Selection | `Presentation.getSelectedSlides()`, `setSelectedSlides(ids)`, `Slide.setSelectedShapes(ids)` | 1.5 |
| Range picture | `Excel.Range.getImage()` -> base64 PNG | ExcelApi 1.9 (a runtime floor, not a manifest one - see the two-hosts row) |
| Chart picture | `Excel.Chart.getImage(w,h,"Fit")` | ExcelApi 1.2 |
| Range anchor that follows edits | hidden workbook-scoped defined name (`NamedItemCollection.add`, `NamedItem.visible=false`, `getRangeOrNullObject`) | ExcelApi 1.4 / 1.1 |
| Chart anchor | `Chart.name` (writable) resolved across all worksheets | ExcelApi 1.1 |
| Excel-side registry | `workbook.settings` (already used by the audit overlay) | ExcelApi 1.4 |
| File name | `Office.context.document.getFilePropertiesAsync()` -> `url` ("" until saved) | Common API |
| One add-in, two hosts | `<Hosts>` with `Workbook` + `Presentation`, per-host `VersionOverrides`; SharedRuntime 1.1 in PowerPoint (Win 2102, Mac 16.46). The XML manifest cannot scope `<Requirements>` per host - a top-level ExcelApi set would hide the add-in in PowerPoint - so the ExcelApi 1.9 floor is enforced at runtime in the Excel pane boot (`src/main.ts`), and the custom functions are gated by their Workbook-only extension point rather than by a requirement set. | manifest |

Unverified (spike, section 10): tags surviving cut/paste across slides and copy into another
deck; `Range.getImage` orientation and pixel density on Mac; whether `OfficeRuntime.storage`
is shared between the Excel and PowerPoint panes of one domain; the team's PowerPoint builds.

Corrections to `docs/research/officejs-feasibility.md`: `Shape.getImageAsBase64` is 1.10 (not
1.8); `ShapeCollection.addPicture` is preview-only (no production image-insert API);
`Range.getImage` is ExcelApi 1.9; `Shape.tags` (1.3) was never mentioned.

## 3. Architecture

```
Excel pane (src/excel/links.ts)                 PowerPoint pane (src/ppt/)
  anchor: hidden name / chart name                 anchor: shape tags SMT_LINK + SMT_KEY
  registry: workbook.settings["SMT_LINKS"]         scan slides -> shapes -> tags
  render: Range.getImage / Chart.getImage          refresh: fill.setImage (or reinsert)
        |  seal (AES-GCM, key from token)                ^  open (same key)
        v                                                |
  PUT /api/links/{id}   POST /api/inbox/{ws}      GET /api/inbox/{ws}  GET /api/links/{id}
        \______________ Rust relay (server/, sqlite, ciphertext only) ______________/
```

Identity is the tag, never shape id / name / slide / position. Geometry is never written on
refresh (except height when the picture's aspect ratio changed), so "moved" and "resized"
need no special handling: the refresh finds the shape wherever it is and repaints its fill.

## 4. Data model (`src/link/model.ts`, pure)

```ts
type LinkId = string;            // 16 random bytes, base64url (22 chars); id8 = first 8 chars
type LinkKind = "range" | "chart";
interface Source { workbook: string; sheet: string; ref: string; anchor: string }
interface LinkTag { v: 1; id: LinkId; kind: LinkKind; rev: number; src: Source; pushedAt: string }
// PowerPoint shape tags: SMT_LINK = JSON.stringify(LinkTag); SMT_KEY = token (base64url, 43 chars)
interface RegistryEntry { id: LinkId; kind: LinkKind; anchor: string; label: string;
  token: string; createdAt: string; lastPushedAt: string | null; rev: number }
// Excel: workbook.settings["SMT_LINKS"] = JSON.stringify({ v: 1, links: RegistryEntry[] })
interface Payload { v: 1; kind: "picture"; mime: "image/png"; width: number; height: number;
  png: string; src: Source; pushedAt: string; hash: string }   // kind "table" reserved for v2.1
interface InboxItem { id: LinkId; token: string; kind: LinkKind; label: string; src: Source; createdAt: string }
```

Anchors: range -> defined name `SMT_LINK_<id8>` (hidden, workbook scope); chart ->
`chart.name = "SMT_LINK_<id8>"`. A `#REF!` name or an unfound chart name = "source missing".

## 5. Crypto (`src/link/crypto.ts`, WebCrypto, pure)

- token = 32 random bytes; `encKey = HKDF-SHA256(token, salt="", info="smt-link-enc")`,
  `authKey = HKDF(token, info="smt-link-auth")`, both 32 bytes, base64url on the wire.
- seal: AES-256-GCM, 12-byte random IV, AAD = link id (UTF-8); blob = IV || ciphertext+tag.
- Workspace key: 32 random bytes; `wsEnc/wsAuth` with infos `smt-ws-enc` / `smt-ws-auth`;
  `wsId = base64url(HKDF(wsKey, "smt-ws-id"))`. Inbox items are sealed with wsEnc, AAD = wsId.
- The server stores `sha256(authKey)` only. It cannot read payloads or forge writes.

## 6. Relay (`server/src/{main,relay,store}.rs`, axum, same origin as the panes)

| Route | Auth | Behaviour |
|---|---|---|
| `PUT /api/links/{id}` | Bearer authKey | create on first PUT (stores sha256), else must match (403); body <= 4 MB (413); rev += 1; keeps last 2 revs; TTL 7 d from last PUT |
| `GET /api/links/{id}` | Bearer | latest blob, `ETag: "<rev>"`, honours `If-None-Match` (304); 404 when unknown/expired |
| `DELETE /api/links/{id}` | Bearer | removes all revs |
| `GET /api/links/{id}?rev=<n>` | Bearer | one named revision (the revert path), never a 304; 404 once retention or the TTL has dropped it |
| `POST /api/links/status` | per item `{id, auth}` | `[{id, rev, pushedAt}]`, unknown -> `rev: null`, wrong auth -> `rev: null, error: "auth"`; max 200 items, so a bigger deck polls in several requests |
| `POST /api/links/fetch` | per item `{id, auth, knownRev?}` | `{items: [{id, rev, blob}], omitted: [{id, reason}]}`; max 200 items and 4 MiB of blobs per answer, the overflow named `deferred` and fetched one GET each |
| `POST /api/inbox/{ws}` | Bearer wsAuth | body = sealed InboxItem <= 64 KB; TTL 24 h |
| `GET /api/inbox/{ws}` | Bearer wsAuth | `[{id, createdAt, blob}]` newest first |
| `DELETE /api/inbox/{ws}/{id}` | Bearer wsAuth | remove one |

Store: sqlite via `rusqlite` (bundled, WAL), path `MODELIS_DATA` (default `./data`). Tables
`links(id, rev, auth_hash, pushed_at, expires_at, blob)` (PK id+rev) and
`inbox(ws, id, auth_hash, created_at, expires_at, blob)`. Sweeper: on every write + hourly
task. Errors are JSON `{error}`; bodies are `application/octet-stream`. The existing
cache-control middleware prefix `/assets/taskpane-` generalises to `/assets/`.
Public path: `dbautomatizacijas.com/modelis/api/...` (gateway strips `/modelis/`; the
Cloudflare Access bypass already covers it). `CLAUDE.md` records the exception: the bypassed
path may serve encrypted relay blobs because they are unreadable without a key that only
exists inside a deck or workbook.

## 7. Flows

Export (Excel): selection or active chart -> new id + token -> anchor created (name / chart
name) -> registry entry -> render (`Range.getImage()`; `Chart.getImage(w*2, h*2, "Fit")`)
-> Payload sealed -> PUT -> InboxItem sealed with the workspace key -> POST inbox -> toast
"Sent to PowerPoint: <label>".

Insert (PowerPoint): Inbox lists pending items -> "Insert" -> rectangle on the selected
slide sized to the picture's aspect (width = min(picture pt width, slide width - 2 x 36 pt),
centred) -> `fill.setImage` -> tags written -> inbox DELETE. Below PowerPointApi 1.8:
`setSelectedDataAsync(Image)` on the active slide, then tag the newest shape.

Push (Excel): "Push selected" / "Push all" re-renders each link through its anchor (never
through the stored address) and PUTs; a `#REF!` anchor marks the entry "source missing" and
is skipped with a visible reason.

Update (PowerPoint): scan every slide's shapes for `SMT_LINK` (one batched load per slide)
-> `POST status` -> per link: up to date (rev equal) / update available / missing (rev null)
/ wrong key (auth error). "Update selected / slide / all" -> GET (If-None-Match) -> open ->
`fill.setImage`; if |newAspect - oldAspect| / oldAspect > 0.5 %, `height = width * h/w`
(top-left and width kept) -> tag rewritten with new rev/pushedAt/src. If `src.workbook`
differs from the tag's, the row shows "source changed: <old> -> <new>" (ambiguous source).
Below 1.8: delete + reinsert at the same left/top/width/height + re-tag (z-order restored
only when `setZOrder` exists).

Break link (PowerPoint): delete both tags; the picture stays. Remove link (Excel): delete the
name (or reset the chart name), drop the registry entry, DELETE on the relay.

Pairing: Excel Settings -> "Link key" (generate, copy); PowerPoint Settings -> paste once.
Stored in `OfficeRuntime.storage` with localStorage fallback behind `WorkspaceKeyStore`. If
the spike shows the storage is shared across hosts, pairing is automatic.

## 8. Errors and limits

Every relay call maps to a typed `RelayError` (`network`, `auth`, `missing`, `tooLarge`,
`server`) rendered by the shared toast with "Copy details". Update-all continues past
per-link failures and reports "12 updated, 1 missing, 1 wrong key". Payload > 4 MB is
refused before upload with the size shown. Selections above the existing
`SELECTION_CELL_CAP` are refused as today. Tag values above 2 KB are a bug, guarded by the
codec. Undo: PowerPoint cannot undo add-in changes; the relay keeps the previous rev so a
"Revert" button is a UI-only addition later.

Security note shown once in the Links list: anyone holding a deck can pull that link's latest
render for up to 7 days; use "Break links" before sending a deck outside.

## 9. Code layout and quality bar

```
src/link/model.ts      types, codecs, validation (pure)          model.test.ts
src/link/crypto.ts     token/HKDF/AES-GCM/base64url (pure)       crypto.test.ts
src/link/relay.ts      RelayClient over fetch, typed errors       relay.test.ts (fetch stub)
src/link/status.ts     status derivation (pure)                   status.test.ts
src/excel/links.ts     Excel adapter                              test/links.integration.test.ts (fakehost)
src/ppt/host.ts        PowerPoint adapter                         test/ppt.integration.test.ts (fakeppt)
src/ppt/main.ts        PowerPoint pane entry (pptpane.html)
src/ui/*               toast, guard, tabs, report (shared)        unit tests
test/fakeppt.ts        fake PowerPoint host, strict load semantics
server/src/relay.rs    routes                                      #[tokio::test] on in-memory sqlite
server/src/store.rs    sqlite store + sweeper
manifest/spec.ts       single source for both hosts and both environments
scripts/build-manifests.ts
```

Prerequisites (Tier 1 of the approved plan): `src/excel.ts` split into `src/excel/*` with
an `index.ts` barrel; manifest generator + `manifest:check`; dev loop restored (`npm start`
links the dev manifest, `poststop` restores the prod copy); one version source +
`version:check`; `npm run check` (tsc, eslint strict, prettier, vitest, cargo test, checks)
in GitHub Actions; `src/ui/report.ts` error surface; shared UI extracted from `main.ts`.

Hard constraint while touching manifests: Daniel's `wef` file is currently a HARD LINK to the
repo's `manifest.xml`. Before the generator ever writes `manifest.xml`, replace the wef entry
with a plain copy of `manifest.prod.xml`; generators write via temp file + rename.

Rules: no `any`, no `console.*`, no TODO; one purpose per file, pure logic without Office.js
imports; Office.js calls only in the two adapters; every new module ships with tests; strict
load semantics in both fake hosts.

## 10. Spike before building on unverified facts (~1 h, real Office on the Mac)

1. Tags: add tags to a shape; save, reopen; cut/paste to another slide; duplicate the slide;
   copy into another deck; move via Selection pane -> tags still present?
2. `addGeometricShape` + `fill.setImage` renders; second `setImage` refreshes; geometry
   unchanged; z-order unchanged.
3. `Range.getImage()` on Mac: orientation (known flip bug), pixel size of a 10x5 range.
4. `OfficeRuntime.storage.setItem` in Excel, `getItem` in PowerPoint (same domain).
5. Team PowerPoint builds -> PowerPointApi level (decides how much the <1.8 fallback matters).
Needs the dev manifest with the PowerPoint host (after the generator lands). Findings go to
`docs/research/officejs-feasibility.md`.

## 11. Tests (definition of done)

Unit: codec round-trips, tag key uppercasing, size guard; seal/open, wrong key fails, AAD
binding; status matrix; RelayClient error mapping (200/304/403/404/413/500/network).
Excel integration: export registers hidden name + registry; render after a simulated row
insert reads through the name; `#REF!` -> source missing; chart resolved by name across
sheets; remove link cleans all three places.
PowerPoint integration: insert -> tags; shape moved to another slide -> refresh finds it,
calls `setImage`, writes no geometry; aspect change writes height only; 404 -> missing;
changed workbook -> warning; break -> tags gone, picture intact; <1.8 fallback path.
Server: put/get round trip, 403, TTL expiry, rev increments, 2-rev retention, 413, inbox
lifecycle, sweeper, status batch.
Real Office pass (~20 min): export one range + one chart; insert; move both to other slides,
resize; insert rows above the source; change numbers; Push; Update all -> refreshed in place
with kept sizes; delete the source rows -> "source missing"; wrong workspace key -> empty
inbox + clear error; deck opened without the add-in -> plain pictures.

## 12. Rollout

v2.0.000 (new host = major). `db deploy modelis` builds both bundles + server; systemd unit
gains `MODELIS_DATA` (the host's data directory). M365 manifest update adds the PowerPoint
host (admin, same screen as the initial upload); JS-only updates need no admin action.
