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
