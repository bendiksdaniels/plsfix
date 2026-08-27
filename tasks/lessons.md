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
