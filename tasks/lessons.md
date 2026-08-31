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
`git worktree list` right after launching.

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
