# Local links: Excel to PowerPoint on one computer, the relay optional - design

Status: design 2026-09-23, approved direction (Daniel, 23.09: "work more on how the linking feature
would work on local device without servers", then "Both: local by default, relay optional", then
"go"). Scope: a second transport for the link feature. Goal: a new install moves linked objects
from Excel to PowerPoint with no relay at all, so no link data leaves the computer; the relay stays
as a switch, and a device that already uses it keeps it.

Not in scope: starting the pane with no network (the pane itself still loads from the host in the
manifest, which Office caches); keeping two computers in sync without a relay; auto-push in local
mode; a native helper app; any change to the relay server or its API.

## 1. What stays the same

Link identity (the `PLSFIX_LINK` / `PLSFIX_KEY` shape tags, the registry in the workbook), the
payload model (`src/link/model.ts`), per-link keys and sealing (`src/link/crypto.ts`), and every
flow in `src/excel/` and `src/ppt/`: export, push, insert, update, revert, change source, paste
latest linked. Both panes already reach the relay only through `RelayApi` (`src/link/relay.ts`):
the local transport is two more implementations of that interface, not a second copy of the flows.

## 2. The carrier: one clipboard copy, two flavors

A **bundle** is everything PowerPoint needs to insert or update a set of links:

```json
{ "plsfix": "links", "v": 1,
  "links": [ { "id": "<32 hex>", "rev": 1099511627781, "sentAt": 1790000000000, "blob": "<base64url>" } ],
  "inbox": [ { "id": "<32 hex>", "createdAt": 1790000000000, "blob": "<base64url>" } ] }
```

- `links[].blob` is the payload sealed with the link's own key, byte for byte what a relay PUT
  would carry. `inbox[].blob` is the inbox item (it holds the link token) sealed with the **local
  workspace**: `deriveWorkspace` of a fixed all-zero secret, the same on every device. That seal
  protects nothing and says so in the code: it exists only so both transports share one inbox
  envelope and `listInbox` / `insertFromInbox` run unchanged. The bundle is therefore a secret
  while it holds inbox rows: anyone with the text can open those links.
- Every copy puts two flavors on the clipboard. `text/html` is one `<span>` whose
  `data-plsfix-links` attribute holds the bundle (base64url of the JSON) and whose text is the
  sentence below. `text/plain` is only the sentence: "pls,fix links for PowerPoint: paste them in
  the pls,fix pane, Inbox tab." A paste onto a slide, a cell or an e-mail gives that sentence, never
  megabytes of text.
- **Writing** (Excel): the click handler starts
  `navigator.clipboard.write([new ClipboardItem({ "text/html": promise, "text/plain": promise })])`
  synchronously, and the promises resolve when the export or copy has rendered. When the engine has
  no `ClipboardItem`, or refuses the write (the pane lost focus, a permission), the pane keeps the
  prepared bundle and shows a **Copy for PowerPoint** button: its press copies synchronously through
  a `copy` event (`execCommand("copy")` setting both flavors). If that fails too, a read-only box
  shows the bundle as plain text with "Select all, then Ctrl+C (⌘C on a Mac)".
- **Reading** (PowerPoint): a paste box on the Inbox tab (an editable `<textarea>`, so every host
  fires `paste` on it). The `paste` handler reads `text/html` first (DOMParser, the
  `[data-plsfix-links]` attribute), then `text/plain` as a bundle (the manual fallback), calls
  `preventDefault`, and never renders the text into the box.
- Caps: a bundle over 25 MB is refused in Excel ("That is too much to copy at once: copy one project
  or the selected links."); every per-payload cap is unchanged.
- Proven 23.09 in headless Chromium (the WebView2 engine), throwaway probe: both write paths carry
  5 MB in the HTML flavor, plain text stays the sentence, the box stays empty, and a one-click write
  whose content arrives 45 s later still lands. Not yet proven: WKWebView in Office for Mac and
  Office on the web (section 9).

## 3. The transport setting, per device, in both panes

- One `<select id="link-transport">` per pane: **Copy and paste on this computer** (`local`) or
  **Through the relay** (`relay`). Stored per device under `plsfix.link.transport.v1` through the
  existing `KeyStore` (`officeKeyStore`).
- Default when nothing is stored: `relay` if this device already holds a link key
  (`plsfix.link.workspace.v1`), else `local`. Existing users keep what works; new installs start
  local.
- Excel, Links tab: a last section "HOW LINKS REACH POWERPOINT" holds the select and one sentence
  per mode. The Link key section and the Auto-push box show only in relay mode (auto-push cannot
  write the clipboard on an edit). In local mode "Push selected" / "Push all" read "Copy selected" /
  "Copy all", every "push" or "sent" word on the tab reads "copy" / "copied", and the export toast
  says "Copied for PowerPoint: <label>. Paste it in PowerPoint: pls,fix, Inbox tab."
- PowerPoint, Settings tab: the same select above the Link key controls, which show only in relay
  mode. Inbox tab in local mode: the paste box ("Paste from Excel: click here, then Ctrl+V (⌘V on a
  Mac)") and "Clear pasted links"; the unpaired hint only in relay mode.
- Switching mode never deletes anything: the stored key, the relay copies and the pasted links all
  stay where they are.

## 4. Excel side: the writer

- `src/link/local-collector.ts`, `LocalCollector implements RelayApi`: `putLink` records
  `{ id, rev, sentAt, blob }` and answers the rev; `postInbox` records the row; `deleteLink` does
  nothing (no relay copy to revoke: the deck keeps its object); `touchLinks` answers the count;
  every read method throws (a programming error on this side). `bundle()` returns what it recorded.
- **Revisions**: local revs live above `LOCAL_REV_BASE = 2 ** 40`; relay revs are small integers
  the server counts. The next local rev of a link is `max(registry rev, LOCAL_REV_BASE) + 1`, read
  from the workbook registry at the start of the action (the collector takes a `revOf(id)` lookup).
  The two spaces never meet, so a switch in either direction still shows as an update (the deck
  compares revs by inequality, `deriveStatus`), and a rollback is still refused by the payload's
  own `pushedAt` (`assertFresh`). Revert's previous rev inside the local space is `rev - 1` down to
  `LOCAL_REV_BASE + 1`.
- Per action: an export copies one link plus its inbox row; **Copy selected** / **Copy all** (the
  Push scope rule: the ticked rows, or the project the list shows) copy those links plus an inbox
  row for each, so a deck that does not hold one yet can insert it. `pushLinks` takes an optional
  `announce?: Workspace` for that, reusing `announce()` from `src/excel/link-record.ts`.
- The registry's `rev` and `lastPushedAt` move on a copy exactly as on a push.
- A copy that fails after the link was created says so: "Copy failed: <label> is linked but not
  copied. Tick it in the list and press Copy selected."
- Where: `src/pane/links-transport.ts` (new) owns the setting, the default rule and
  `beginCopy(action)` (the synchronous clipboard start, the prepared-bundle fallback and the manual
  box); `src/ui/clipboard-links.ts` (new) owns the two-flavor write and read. `links-tab.ts` only
  switches its call sites and labels.

## 5. PowerPoint side: the reader

- `src/link/local-store.ts`, `LocalStore implements RelayApi`, holding per link the two highest
  revisions it has seen (the relay keeps two as well) and the inbox rows:
  - `status`: a known link answers `{ rev, pushedAt: sentAt }`; an unknown one answers
    `{ rev: null, pushedAt: null, local: true }`.
  - `fetchLinks`: rev equal to `knownRev` is omitted `unchanged`; an unknown link is omitted
    `notPasted`; no `deferred`, no batch cap.
  - `getLink` / `getLinkRev` answer an exact revision or throw `RelayError("missing")`.
  - `listInbox` / `deleteInbox` work on the local-workspace rows.
  - `ingest(bundle, deckIds)`: keeps the two highest revs per link (an old bundle never downgrades
    anything), stores inbox rows only for ids the deck does not hold, and returns the ids it took.
- Persistence, `src/link/local-persist.ts`: IndexedDB database `plsfix-links` (stores `links`,
  `inbox`), write-through, read into memory at boot; a 32 MB budget on blob bytes, evicting the
  least recently pasted links first (never the links of the paste in progress). No IndexedDB (a
  private window, a refused open): memory only, and the Inbox says once "This computer does not keep
  pasted links between sessions: paste again after reopening the pane."
- Statuses: `LinkStatus` gains `notPasted`, badge "Not pasted yet", filter option "Not pasted yet".
  `deriveStatus` maps a local answer with no rev to `notPasted`, and a local rev below the tag's to
  `current` (the deck already holds something newer than this computer's copy). `FetchOutcome`,
  `UpdateSummary` and `summarize` count `notPasted` as "N not pasted yet".
- **Paste flow**, `src/ppt/paste-links.ts` (new), through `act()` so the busy latch holds: decode,
  scan the deck's link ids, `ingest`, then `updateLinks` on the deck rows whose id came in the bundle,
  then re-read the Inbox. Toast: "Pasted 12 links: 9 updated, 3 up to date, 2 waiting in the Inbox."
  (parts with zero left out; failures ride in Copy details as on Update all).
- Revert works while the store holds the revision below the tag's. The two relay-worded sentences a
  local run can reach become neutral: "revert <label>: the previous version is no longer available."
  and "The update is older than the picture this deck already holds."
- The Links list column "Pushed" becomes "Sent" (both a push and a copy send).
- "Clear pasted links" (two-click confirm, `src/ui/confirm.ts`) empties the store and the Inbox.

## 6. Security and privacy

In local mode no request carries link data; the pane still loads its own code from the manifest host.
A bundle holds the tokens of the links it carries, so it is meant to be pasted, not forwarded (help
sentence and privacy page). IndexedDB keeps payloads and inbox rows on this device until "Clear pasted
links" or eviction. `public/privacy.html` gains the local store; the relay paragraph stays.

## 7. Errors and edge cases

| Case | What the pane says or does |
|---|---|
| Paste of text that is not a bundle | "That is not a pls,fix copy from Excel." |
| A bundle from a newer add-in (`v` > 1) | "This copy comes from a newer pls,fix. Update the add-in." |
| One blob that will not open | that link fails in the summary; the rest apply |
| Paste while an action runs | the busy sentence, nothing ingested |
| An old bundle pasted after a newer one | nothing downgrades; its links count as up to date |
| IndexedDB full or refused | memory only, the one-time Inbox sentence |
| Relay mode on a host with no relay | the existing relay sentences |
| Clock going backwards between two copies | revs still rise (registry counter); a payload whose `pushedAt` is older than the tag's is refused as today |

## 8. Testing

- Unit: the bundle codec (round trip, guards, version, caps, the HTML carrier, the plain fallback);
  `LocalCollector` revs across the two spaces; `LocalStore` against the FakeRelay rules that apply
  (two revisions, unchanged by equality, exact `getLinkRev`, inbox delete) as one shared contract
  suite run on both; persistence with a fake IndexedDB and with none.
- Integration over the fakes: Excel exports and copies in local mode produce the expected bundles;
  PowerPoint pastes insert, update, revert, change source and report statuses; a link that crossed
  transports (relay rev 3, then a local copy, then the relay again) updates every time.
- Both click-through suites run again in local mode (every button, the paste box, the new select).
- Browser gates: `npm run ux:sweep` and `npm run ux:check` cover the new controls; the throwaway
  probe of 23.09 becomes `scripts/ux/clipboard-probe.mjs` (Chromium) so the carrier stays proven.
- Real Office (Daniel's grant, or his own two minutes): on the Mac, Excel to PowerPoint with a
  one-link export and a 30-link Copy all; paste on a slide gives the sentence; Windows later.

## 9. Open risk and its answer

WKWebView (Office for Mac) and Office on the web may refuse the one-click write or hide the HTML
flavor. The design already degrades in two steps (the Copy for PowerPoint button, then the manual
box), and the reader accepts a plain-text bundle, so the feature works everywhere; the real-Office
check decides how often the extra press shows.

## 10. Rollout

Slices (each reviewed, gated and merged on its own; the patch bumps per merge, the last one a minor):
S1 core (`src/link/`: bundle, collector, store, persistence, statuses; no visible change); S2 Excel
pane and S3 PowerPoint pane in parallel after S1; S4 help copy, manual (Latvian), README, INSTALL,
privacy page, the clipboard probe and the click-through runs in local mode. Main stays shippable:
S2 and S3 merge back to back and the release after them is v2.9.0, then the deploy.
