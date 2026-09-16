# Lessons

## 2026-08-27: the `hidden` attribute is defeated by any author `display` rule

`.view { display: grid }` silently killed `element.hidden = true` (UA-origin `[hidden]{display:none}`
loses to any author-origin display declaration, regardless of specificity). Tab switching shipped
broken; tests and build were green because neither sees CSS cascade.

Rules:
- Every stylesheet in this repo carries `[hidden] { display: none !important; }` in the reset.
- A UI feature is not "done" on tests+build alone: render the pane (browser dev mode or sideload)
  and click the new interaction before claiming completion. The shared Playwright browser can be
  held by another session — fall back to manual sideload confirmation and say so.
- When hiding a container of interactive children, also clear its children (stale listeners on
  visible-again nodes re-apply old state).

## 2026-08-27: agent worktrees live inside the repo

`.claude/worktrees/` sits under the repo root, so a bare `git add -A` embeds a RUNNING agent
worktree as a gitlink, and vitest double-counts its test copies. Fixed durably: `.claude/` is
gitignored and vitest.config.ts pins `include: ["src/**/*.test.ts"]`. Rule: never `git add -A`
while an agent worktree exists unless `.claude/` is ignored (it now is).

## 2026-08-27: never exact-match state Excel gives back

Excel rewrites number formats on read-back (currency symbols become locale-tagged codes like
`[$€-x-euro2]`; Mac collapses `hh:mm` to `h:mm`). Any cycle/toggle that derives its position
from read-back state must canonicalize both sides before comparing (cycles.ts
canonicalNumberFormat) — and edge borders (`edgeTop`/`edgeBottom`) style the RANGE, not each
row: per-row semantics need a per-row loop.

## 2026-08-27: the package.json `config` block for Office add-ins is load-bearing

`office-addin-debugging start` only WAITS for the dev server when `config.dev_server_port`
exists in package.json; without it, vite is spawned detached and Excel is sideloaded
immediately — a cold start opens a dead pane URL ("Sorry, we can't load the add-in") that
works on retry, indistinguishable from a crash to the user. Also: Node 17+ resolves
`localhost` to `::1` first, so vite binds IPv6-only and IPv4 probes get refused —
`dns.setDefaultResultOrder("ipv4first")` in vite.config.ts is part of the template for a
reason. Debug shortcuts learned: a missing
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` folder means sideload never
registered (failure is BEFORE Excel); `npm start -- --no-sideload` runs the whole tooling
pipeline headlessly; the strict-load fake host + deleting `.load()` calls one at a time
(mutation sweep) proves load-ordering correctness without touching Excel.

## 2026-08-29: agent worktrees fork from the session-start HEAD, not from main

Every `isolation: worktree` agent started at the commit the session opened on (9e4b7f4), five
commits behind main; three of them noticed, one built on the stale tree. Rule: the first line
of every dispatch is `git reset --hard <current main sha>` and the report must name that sha.
Merging a branch whose conflict hunk cuts through a CSS rule must be resolved from the two
full file versions (main + the branch's appended block), never by keeping both hunk sides.

## 2026-08-29: fix rounds cost more than the build

A 6-line fix through a resumed implementer cost 313k tokens (it re-verified everything).
Rules now in the ledgers: fixes under ~10 lines with an unambiguous spec are applied by the
controller and covered by the named tests; doc/script-only tasks and pure-logic tasks whose
tests were fully specified in the plan get a controller review, opus reviews go to adapters,
crypto and the server. One review pass per task; scoped re-reviews only when a gate fails.

## 2026-08-29: Office on the web is the verification rig when the desktop is off limits

No desktop control (Daniel will not grant computer-use). What worked: a scratch Google Chrome
with `--remote-debugging-port` driven by `playwright-core` over CDP, Office for the web with
the add-in registered through the document URL (`wdaddindevserverport`, `wdaddinmanifestfile`,
`wdaddinmanifestguid`) from a local HTTPS+CORS manifest server on the office-addin-dev-certs;
Chrome needs `--ignore-certificate-errors-spki-list=<cert spki>` and the local-network-access
check off, the unified Apps store has no "Upload My Add-in" any more, and pre-existing
cross-origin frames report an empty `url()` on a fresh CDP connection (match frames by
`location.href`). Every pane action is then verifiable through `Excel.run` from inside the
pane frame instead of screenshots.

Three things the web pass found that tests and the desktop had not: `Office.onReady` never
settles when the custom-functions runtime fails to initialise (pane stuck on "Connecting",
ribbon commands dead: `src/host-ready.ts` now probes the host after a head start); Excel for
the web rejects `chart.format.font` and `roundedCorners` on chartex charts (the waterfall
applies its surface in a tolerated batch of its own); and "Fill formula right" sizes by the
neighbouring rows, so a demo row with blank neighbours cannot be filled. Deleting rows inside
an exported range shrinks its hidden name; only removing the whole block gives "source missing".

## 2026-08-30: a worktree agent forks from the shell's CURRENT directory, not the project

Three agents launched right after a `cd` into the memory repo got worktrees of the memory repo,
found no source and stopped (about 180k tokens for nothing). Rule: run `cd ~/plsfix && pwd` as
the last shell command before any `isolation: worktree` dispatch, tell the agent which repo it
must be in (`git remote -v`, `ls package.json`) and to report BLOCKED otherwise, and verify with
`git worktree list` right after launching. Repeated 01.09: a `cd ~/claude-vm` in a Bash call sent
in the SAME message as two dispatches made both fail ("Failed to resolve base branch HEAD").
The dispatch message carries the Agent calls and nothing else.

## 2026-08-30: an agent symlinked the repo's node_modules into its worktree

`~/plsfix/node_modules` became a symlink to a worktree's install; the gates then died with
exit 194 while `tsc` still ran. Rule for every dispatch brief: never touch the main checkout's
`node_modules` (worktrees get their own `npm install`); the controller checks
`test -d node_modules && ! test -L node_modules` before running the gates after a merge.

## 2026-08-30: Office on the web hangs, jams and heals in ways the tests cannot show

Native-chart spike on PowerPoint for the web (CDP rig, `scratchpad/driver`): `PowerPoint.run`
never returns while the PowerPoint tab is in the background (`page.bringToFront()` first,
always); a batch that never returns (a 48-shape insert, a `slides.add()`) jams every later
WRITE while trivial reads keep answering, which looks like "the selection API is broken" -
the fix is a page reload (the `wdaddin*` parameters stay in the URL and the add-in
re-registers without the dialog; reopen the pane through the ribbon: tab "pls,fix", button
"Links"); Office Online can answer one navigation with "services aren't available right now"
and the next with the deck. Cost model on the web: an add costs more the more shapes the slide
already holds (12 rectangles 1.7 s on a clean slide, 24 in 15 s at 45 shapes), a text box
about 1 s whatever its property count, so shape-built charts need syncs of ~12 shapes and a
per-host shape budget. Pie adjustments are degrees clockwise from 3 o'clock, normalised to
(-180, 180], addressable only after the shape's first sync. Spike scripts:
`spike-*.js` in this session's scratchpad; findings in
`docs/superpowers/specs/2026-08-30-native-charts-design.md`.

## 2026-08-30 late: a web proof must keep the pane's host tab in front until the pane is idle

The first native-chart proof "failed": the pie's 21 shapes were on the slide, ungrouped and
untagged, the pane stuck busy. The adapter was fine (a faithful replay of its writes grouped
21 shapes in 13 s); the script had brought the Excel tab to the front three seconds after
clicking Insert, and a `PowerPoint.run` in a background tab never returns. Rules for the rig:
wait for the pane's own busy flag (`#tab-inbox` disabled) plus a toast change, never a fixed
delay; switch tabs only while the pane is idle; a pane stuck busy is reset by reloading its
frame (`location.reload()` inside the pane), not the deck. Measured on the web: a 21-shape pie
18 s, a 20-shape column chart 24 s, "Update all" over three chart links 54 s.

## 2026-08-31: never drive Daniel's desktop Office, not even by sideloading

A probe agent sideloaded a dev manifest into Daniel's running Excel (`office-addin-debugging
start ... desktop`) to reproduce a desktop-only export failure. Daniel: "Stop trying to open
excel or simulate the environment." Rule: desktop Excel and PowerPoint on his Mac are HIS;
no sideload, no `npm start`, no AppleScript, no manifest swap in `wef/` unless he runs it
himself. Real-Office evidence comes from him (the pane's "Copy details" block carries the
Office error `code` and `debugInfo`), and automated verification stays on Office for the web
through `scripts/rig/` (a scratch Chrome he signs into). After any such stop: kill the
tooling, restore the prod manifest (`scripts/wef-restore-prod.sh`), verify both `wef/`
folders, drop the branch.

## 2026-09-08: `npm start` without `npm stop` left the desktop on localhost for a week

Codex's 01.09 `npm start` / `start:ppt` put the DEV manifest (v2.4.9, `https://localhost:3000`)
into both wef folders and left two vite processes alive for six days; Daniel's desktop add-in
worked only through that orphan, with the old ribbon, while production sat unused. Rules:
- A sideload session ends with `npm stop` (its poststop runs `scripts/wef-restore-prod.sh`);
  if the tooling is gone, run the restore script by hand and `pgrep -fl vite` for leftovers.
- Proof of the restore is `cmp` of each wef manifest against `manifest.prod.xml`, and
  `lsof -nP -iTCP:3000 -sTCP:LISTEN` empty. Office reloads the manifest only on relaunch.

## 2026-09-08: never fold an `await` into an argument list after a `.value` read

`picture(image.value, await readChartData(...))` reads the client result BEFORE the awaited
call commits the batch that loads it: the strict fake threw ValueNotLoaded on the sharp path,
the plain retry ran, and a "no extra round trip" test counted one more sync. The two-statement
form (`const read = await ...; return picture(image.value, read);`) is the rule for office.js
code: the sync first, the `.value` after, never inside one expression.

## 2026-09-08: the web rig, three things that were not true on 30.08 any more

- Office on the web's "Enable Developer Mode" dialog now carries an opt-in checkbox (Excel
  `#WACDialogOptInCheckbox-input`, PowerPoint `#optInCheckbox`); OK without the tick only
  dismisses it, and the tick counts on the NEXT load: reload the same URL and answer the
  "Registering Developer Add-in Manifest" Yes there (`scripts/rig/snippets/register.js`).
- `prettier` had put a `;` after every snippet's arrow function, which `drive.mjs` wrapped
  into `(fn;)`: every `@file` run died with "Unexpected token ';'". The driver strips it now.
- The manifest server logs every request; run it detached with a log file, never behind
  `| head`: the first write after the reader closes kills it and Office reports nothing.
- A PowerPoint-for-the-web write batch can hang without rejecting (a 21-shape pie stopped
  after its first chunk of 12 shapes); the pane stays "busy" for ever and the twelve loose
  shapes stay on the slide. Heal: `location.reload()` in the pane frame, delete the untagged
  leftovers, insert again (20 s the second time). Poll busy + toast + shape count so a jam
  shows within a minute instead of a 210 s silent wait.

## 2026-09-09: a ten-agent audit fleet, what the run itself taught

- The account's session limit stops every agent at once (all ten at 22:xx, reset 23:30).
  An agent with a commit resumes through SendMessage with its worktree intact; an agent
  with no commit loses its worktree (auto-cleaned as unchanged) and has to be dispatched
  again from scratch. Brief "commit early", and count the loss window in the estimate.
- Subagents get the harness's own attribution reminder and add `Co-Authored-By` and
  `Claude-Session` trailers despite the brief (three of ten branches). Strip them on a merge
  branch first: `git filter-branch -f --msg-filter 'sed "/^Co-Authored-By: Claude/d;
  /^Claude-Session:/d"' <base>..merge-x`; the tree stays byte-identical.
- Merge branches are rewritten and rebased, so `git branch --merged` never lists the agent
  branch; `git cherry main <branch>` with no `+` line is the merged test before a worktree
  is removed.
- One slice's rule breaks another's test at the rebase, never before: B's "the overlay owns
  only fills it striped" failed E1's test that toggled the overlay over a formula-free
  block (seed two adjacent formulas with one R1C1 form). Run the full check on the rebased
  branch, not on the agent's own base.
- A reviewer's "missed" item is a claim, not a fact: E1's reviewer said a table push below
  ExcelApi 1.9 still throws raw; the try/catch in `renderTable` already falls back to the
  plain grid. Read the code before applying a review fix.
- The fake PowerPoint applies a repaint's `fill.setImage` and `tags.add` to an existing
  shape at once, before the sync that carries them, so a hung sync rolls back a brand-new
  add and nothing else; and shapes seeded by a fixture stay pending until a successful sync
  confirms them, so a test that hangs the very first round trip loses its fixture. One
  ordinary round trip first, then arm `hangNextSync`.
- `ppt.charts.audit` "fails rather than waits when the placement read is the swallowed one"
  tripped vitest's 5 s timeout once inside a full `npm run check` while two agents ran their
  own suites on the machine; it passes alone. If it recurs, give its settle loop fewer steps
  or the test a longer timeout, never a retry.

## 2026-09-09: the web rig at v2.6.14, what the new deadline showed

- The deadline is a symptom detector, not a cure: with every PowerPoint sync under 60 s,
  Update all on the web took 256 s and four batches "stopped answering" - both table
  repaints and both chart-group redraws. Timing the host by hand explained it: a formatted
  table cell cost about 0.4 s per property write on PowerPoint for the web that day (24 cells
  text-only 0.7 s, text + bold + fill + alignment 28 s), so a whole 6x4 repaint was about 120
  writes and 60 s in one batch. Fix: `CELLS_PER_SYNC` 8 in `src/ppt/tables.ts`. Same proof
  afterwards: 128 s, 6 updated, 2 up to date, no deadline anywhere.
- An abandoned batch does not stop: the host keeps working the timed-out request and every
  later batch queues behind it, so one slow table turned the chart redraws after it into
  deadline hits too (the cascade). Size the batches so no healthy one can reach the deadline;
  a hit means a batch is too big, not that the constant is too tight.
- The rig drives the host with its tab in front (`page.bringToFront()`): a background tab's
  `PowerPoint.run` never returns, and the pane frame's `document.visibilityState` is no
  usable signal for it.
- Exports by chart name failed on the rig because the 08.09 session had renamed the demo
  charts to their `PLSFIX_LINK_` anchor names; the picker now lists those names. Not a
  regression, a UX call for Daniel (E1 flagged it) - and a snippet that names a chart must
  read the picker first.
- The link key is a secret: the rig copied it by intercepting `navigator.clipboard.writeText`
  under the Copy button, kept it in the session scratchpad only and deleted it at teardown.
  Never print it, never commit it.

## 2026-09-12: a fake-clock settle loop must yield real event-loop turns

The five "host never answers" suites (test/j.*.audit, ppt.charts.audit) each carried a
copy of `settle`: ten `vi.advanceTimersByTimeAsync(SYNC_TIMEOUT_MS)` steps, no real
pause. Green on the Mac (3-40 ms per test), red on three GitHub runs in a row (5 s, then
20 s timeouts): the decrypt or fetch before the hung sync runs on the thread pool, and on a
slow runner it finished only after the loop had run dry, so the deadline timer it then
armed was never advanced and the promise never settled. Proven by delaying WebCrypto 40 ms:
old loop hangs every test, the fixed helper passes. Rules:
- One helper, `test/hung-sync.ts` `settleHungSync`: a real 10 ms pause per step, up to
  200 steps. Never copy a clock-driven loop into a suite again.
- A CI-only timeout on a fake-timer test is a missing real turn, not a slow test: a longer
  `testTimeout` changes nothing (the 09.09 note said "a longer timeout"; that was wrong).
- Reproduce the runner locally by slowing the async work (a spy with a real delay), not by
  loading the CPU: ten `yes` processes and `UV_THREADPOOL_SIZE=1` did not reproduce it.

## 2026-09-13: the v2.7 wave (ten agents, two sessions on one checkout)

- Two slices that both APPEND to the same file conflict at the rebase even when they append to
  different sections (taskpane.html, dispatch.ts, dispatch.test.ts imports, copy.excel.ts,
  excel/index.ts): keep both sides, HEAD first; an HTML block cut by the marker needs its own
  closing tags before the other block; an import list becomes the sorted union; a function
  boundary cut by the marker needs its closing brace back before prettier runs. Budget one
  conflict commit per feature slice.
- Agents that do not touch CLAUDE.md / README / FEATURES / tasks (the report carries the Map,
  FEATURES and launch-check blocks) removed the docs conflict every 09.09 merge had.
- `npm run ux:check` and `ux:sweep` default to ports 3131/3132, shared by every worktree: two
  gate runs at once answer ERR_CONNECTION_REFUSED or test the other worktree's build. Every run
  gets its own `--port`, and the brief names the pair per agent.
- The account session limit kills every agent at once, mid-edit. A resumed agent (SendMessage
  to its id) keeps its worktree INCLUDING uncommitted files; brief "commit after every green
  test file" anyway. A merge helper must remove its scratch worktree before checking the
  rebased branch out elsewhere (a branch cannot be checked out twice).
- The merge helper leaves the checkout on `merge-<branch>`: `release.sh` runs only after
  `git checkout main && git merge --ff-only merge-<branch>`. Run from the merge branch it tags
  a commit main does not have yet, the hook pushes the tag alone, and main waits for a push
  by hand (v2.7.10).
- Several agents writing gate logs into the one session scratchpad clobber each other; log into
  the worktree.
- A reviewer's "the validator says X" needs the tool's whole output pasted: the store validator's
  "unreachable" lines meant "not deployed yet", not Cloudflare.
- `git cherry` reports a conflict-resolved commit as unmerged (new patch-id): after a resolved
  rebase, the merged proof is the rebased branch's log, not cherry.
- The CI `check` workflow flaked three times on a fake-clock settle loop (hung-sync tests); the
  last copy is on `settleHungSync` since v2.7.1; the workflow is disabled on Daniel's word and
  re-enabling it is his call.

## 2026-09-13: ownership is settled, never raise it

- pls,fix and dbautomatizacijas.com are Daniel's own: the app and the hostname. A launch or
  monetisation discussion never gets an "IP of intern work" or "the suite's hostname as the public
  face" caveat again. The business seam (hosted relay, licence key) is the only thing to discuss.

## 2026-09-13: a tag run that dies at `gh release create` is re-run by dispatch, not by rerun

- GitHub answered the v2.7.14 push run's `gh release create` with an HTTP 500 (every earlier step
  green, image pushed); `gh run rerun --failed` then ended `startup_failure` with zero jobs. The
  path that works: `gh workflow run release.yml -R bendiksdaniels/plsfix --ref main -f tag=v2.7.14`
  (the workflow checks out the tag itself). Probe before mutating: the notes generator
  (`POST .../releases/generate-notes`) and the tag ref proved our side was fine.
- `gh` reads the repo from the shell's cwd: after a `cd` into another repo, `gh release view` says
  "release not found" and `gh run list` reads the wrong repo. Always pass `-R` here.

## 2026-09-13: the history rewrite, and three shell traps met on the way

- Rewrite recipe: a literal `old==>new` map (longest phrase first, identifiers to today's names, a
  catch-all last), `git filter-repo --force --replace-text map --replace-message map
  --path-rename old:new`, then the proofs BEFORE any push: HEAD tree hash unchanged, commit and tag
  counts unchanged, `git grep -i <name> $(git rev-list --all)`, `git log -p --all | grep -ci`,
  `git log --all --format=%B | grep -ci` all zero, the catch-all delta zero. Office files are zips:
  grep cannot see inside; unzip every historical blob and strip the dirty ids.
- zsh: never name a variable `path` (or `PATH` in lowercase forms): it is tied to `$PATH`, and
  `path=$2` with an empty `$2` makes every later command "not found".
- A `while ...; do grep -q ... && {...}; done || exit 1` fails on an EMPTY list: the loop's status
  is the last grep's, 1 on no match. Test the list explicitly.
- Chrome `--restore-last-session` on an expired Office session restores the sign-in page, not the
  documents; the document URLs survive in the profile's `Sessions/` files (`strings | grep sharepoint`).

## 2026-09-13: the v2.8 wave, what the web proof taught

- SharePoint REST: `ListItemAllFields` answers the list item's GUID, not the file's; a `Doc.aspx?sourcedoc=`
  link built from it says "Item does not exist". The file id is `UniqueId` from the folder's `/Files`
  listing (`GetFolderByServerRelativeUrl(...)/Files?$select=Name,UniqueId`).
- A tab opened through the DevTools `/json/new` endpoint has its cross-origin frames unattached until it
  navigates once: `ribbonTabs` reads the ribbon but `clickIn` times out and `body.innerText` is empty.
  Reload the tab (the `wdaddin*` registration parameters survive) before driving it.
- The driver's `hostPage` picks the host tab with the most frames: close this morning's document tabs
  before registering on a new one, or the registration lands on the old deck.
- An explicit spot is the user's choice: `overlapping` must stay false there, or every insert into a
  guided spot prints "no free space" (v2.8.1).
- The export snippet's toast reads lag one export behind on the web; read the sequence, not each line.

## 2026-09-13 evening: the desktop pass on Daniel's Mac (computer use, granted for Excel and PowerPoint)

- PowerPoint for Mac 16.107 (macOS 26.6) crashes on a shape chart: the chunk batch and the group batch both
  come back (25-28 s and 9-10 s, `RichApi Batch Info`), the last invocation logged is a paragraph alignment
  or a shape id read, then the app dies on its own repaint and relaunches with the deck as [Autosaved]. No
  crash report lands in DiagnosticReports (Microsoft Error Reporting takes it). The table route and the
  picture route (fill.setImage) are fine on the same host. v2.8.2 keeps every chart a picture there
  (`hostDrawsCharts`); the bisection (E1 = no group) is written up in AUTORESUME and waits for the Mac.
- A desktop sideload on the Mac shows no ribbon tab on its own: Home > Add-ins > pls,fix (Developer
  Add-ins) loads it once per launch. After a crash relaunch the pane is gone and the pairing key can be
  gone with it (WebKit storage not flushed).
- `office-addin-debugging start` hard-links `manifest.xml` into BOTH wef folders and fails with EEXIST when
  a file is already there: delete the wef manifests first. Never `cp` over those links afterwards (it would
  overwrite the repo's manifest.xml through the link): delete, then `scripts/wef-restore-prod.sh`.
- The dev pane's `/api` proxies to 127.0.0.1:8804; for a desktop test against the live relay point the
  vite proxy at `https://dbautomatizacijas.com/modelis` (uncommitted, changeOrigin) and the inbox fills.
- PowerPoint AppleScript queries hang (AppleEvent timed out) while the add-in pane is open; `out` is a
  reserved word in its dictionary; `open -a "Microsoft PowerPoint"` activates it safely. Excel AppleScript
  (`select range`, `activate object worksheet`) is the reliable way to set a selection for the pane.
- Computer use: clicking the active ribbon tab collapses the ribbon; a Safari `activate` from a script
  dismisses an open Excel popover; the batch tool's scroll takes `scroll_direction` and `scroll_amount`.
- Safari is read-only for computer use, but System Events reaches its page through the accessibility tree:
  walk `UI elements` recursively (buttons carry the label in `name` or `description`), `click` the element
  or `click at` its centre, `keystroke` after `activate`. GitHub's device page is eight one-char fields
  (type the code without the dash) and its Authorize button stays disabled until a focus or visibility
  event fires with the page focused and the button fully in the viewport: switch tabs and back, then press.
- `gh auth refresh` under expect: reply to the colour query (`ESC]11;?ESC\`) and the cursor query
  (`ESC[6n` -> `ESC[24;80R`) or the prompt never reads the answer; the code is bold, so match
  `[A-Z0-9]{4}-[A-Z0-9]{4}` alone; it expires after 15 min; "Confirm access" (sudo mode, 2FA) is the
  user's own step and cannot be automated.
- A foreign document in front of the app you are driving means the user is at the machine: stop, revert
  what the desktop still runs (dev manifest, experiment edits), say so, and wait for the go.

## 2026-09-13 night: the Mac chart crash, bisected on Daniel's Mac (computer use)

- The crash is `ShapeCollection.addGroup` with a whole chart: 19-20 shapes in one group killed
  PowerPoint for Mac 16.107 every time (E2 same context, E3 fresh context, E5 without the lines,
  E7 rectangles-with-text instead of text boxes, E10/E11 text written after grouping, E13 text-less
  shapes), six filled shapes grouped fine (E4), the same 20 shapes grouped through the ribbon
  (Cmd+Alt+G) grouped fine, and sub-groups of six then one group of the sub-groups (E14a/b, E15 =
  the shipping code) survived insert and an Update all redraw. Not the kind of shape, the count.
- Each experiment cost one crash-relaunch cycle (~4 min): PowerPoint comes back by itself with the
  deck as [Autosaved], the pane closed and, sometimes, the pairing key gone; the inbox item is not
  consumed by a crashed insert, so it can be inserted again. A vite HMR reload of the dev pane
  also drops the pairing sometimes and always returns to the Links tab.
- "Update all" is the cheapest way to test a chart draw variant: no Excel round trip, the redraw
  runs the same drawGroup. It only redraws once the source has a new revision: change a cell in
  Excel and "Push selected" first.
- A re-export of a chart that already carries a link re-pushes that link and posts nothing to
  the inbox; rename the chart object in Excel (AppleScript `set name of chart object 1`) and export
  again to get a fresh inbox item.
- Terminal steals the front every few seconds while Daniel reads the session; a computer-use click
  then fails ("Terminal is not in the allowed applications and is currently in front"). A
  background loop of `osascript -e 'tell application "Microsoft PowerPoint" to activate'` every
  0.7 s for the length of the batch keeps the target in front; `open_application` alone does not.
- Do not press Delete after a click on the slide without checking what got selected: a click on
  the chart's bars selected the deck's dashed rectangle underneath and Delete removed it.

## 2026-09-14: the Mac chart tiers, second round (computer use, stopped by Daniel; written up 16.09)

- `ShapeCollection.addGroup` of ONE shape is InvalidArgument on PowerPoint for Mac 16.107: a 25-shape chart
  tiered 6/6/6/6/1 died on the last chunk. A remainder of one stays ungrouped and rides into the next level
  (`tierUp`, v2.8.9); the fake refuses fewer than two ids the same way.
- A refused batch is not a rolled-back batch: the adds queued BEFORE the refused call stay on the slide and
  their ids never reach the pane, so an id-based cleanup cannot see them (four orphan sub-groups, no top
  group, no link, no toast). Name every primitive and sub-group after its chart in the batch that adds it and
  sweep the slide's top level by that prefix after the id pass (`sweepByName`). `helpers.refuseNextSync`
  reproduces the shape in the fake: it applies, then rejects.
- "Does not work" with no toast and no crash = a refused sync whose cleanup could not find what it left
  behind: look at the slide for untagged shapes named for the chart before suspecting the relay.
- A session that ships five patches and writes no ledger costs the next one an hour of archaeology. The
  commit bodies (`git log --format='%b'`) are the first stop, the transcript under
  `~/.claude/projects/<cwd>/` the fallback; the ledger line is written with the patch, not at the end.

## 2026-09-16: a desktop pass without the computer-use grant, and what the Mac hosts really do

- When the Claude app loses Screen Recording mid-session and the user cannot restart his terminals, the
  Terminal's own permissions still carry a full rig: `screencapture -x` (downscaled with `sips -Z 1440`
  for reading), a 40-line Swift tool posting `CGEvent` mouse, scroll and keyboard events at logical points
  (`scratchpad/desk/click.swift`), `open -a` to bring the app forward (an AppleScript `activate` HANGS
  while the add-in pane is open), AppleScript for Excel selections and cell reads. `text of range` is ""
  for an error cell: read `value` or look at the screen. A `type` lands wherever the focus is: click the
  cell first, the pane steals focus after every button.
- Excel for Mac returns an unfilled cell's fill from `getCellProperties` as `{pattern: null, patternColor: "",
  color: "#FFFFFF"}` (plus `@odata.type` annotations) and `setCellProperties` refuses that shape with
  InvalidArgument at `Range.setCellProperties`; `{pattern: "None"}` alone is accepted. Never write back what a
  read gave you; rebuild the settable object field by field.
- A PowerPoint table created through `shapes.addTable` has no table style at all and every
  `styleSettings.load` on it is refused with GeneralException until a style has been written blind; the
  flags only take once the style exists. Write `style`, then the flags, never read first.
- Custom functions declared on a SHARED runtime must be associated by the runtime page: the
  `<Script><SourceLocation>` under the CustomFunctions extension point is ignored there. A page that does not
  load `functions.js` gives #VALUE! for every function on every desktop and nobody notices until someone types
  one. Excel stores such a formula as `_xldudf_PLSFIX_CAGR(...)`; a workbook that carries the un-prefixed text
  shows #NAME? for ever, add-in or not.
- Cloudflare's browser TTL rewrites `no-cache` to `max-age=14400` for `.js` (functions.js): a runtime change
  can take four hours to reach a webview that already has the file. Excel restores the previous pane session
  after a manifest swap and shows "Add-in Error: this add-in is no longer available": close it, load from
  Add-ins again.
- A sonnet agent reads "vitest was N" from the brief and measures a different N on its branch because the
  base moved: give it the base commit, not a count, and let it measure.

- A shared runtime with `lifetime="long"` keeps the OLD pane code alive across "close the pane, load it
  again from Add-ins": a deploy reaches a running desktop only after the app quits. The pane footer's
  version (or Copy details' `version:` line) says which build is really running.
- Quitting Excel through the menu with unsaved changes raises a save sheet; a scripted "Don't Save"
  click that misses SAVES, and the stale workbook lands over a file just rebuilt on disk. Close the
  workbook with `close active workbook saving no` first, then quit.
- A repaint only redraws a chart on a new revision: to see a drawing fix on an existing link either push a
  change to its source or re-insert it; an item whose source chart is gone from the workbook can only be
  re-inserted from the inbox.

