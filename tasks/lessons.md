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
