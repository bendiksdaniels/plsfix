# AUTORESUME - Model Tools (pinned 2026-08-27 evening)

## State: v1.0.1, feature-complete, verified in a simulated host, NOT yet launched

- Repo `~/plsfix` (Desktop symlink), git local-only, HEAD 815671b, tree clean.
- All 8 v1 chunks shipped (see tasks/v1-plan.md for per-chunk detail): shared runtime +
  ribbon + 34 shortcuts, format cycles, autocolor v2 + color key, audit overlay + Smart
  Track, SMT Undo + paste suite + fast fill, native waterfall + chart tools, Workbook tab
  (TOC, sheet explorer, name scrubber), launch polish.
- Verification: 212 vitest green = 102 pure-logic + 110 end-to-end against
  `test/fakehost.ts` (in-memory Office.js host, all 38 excel.ts exports, host-quirk
  switches). 5 review passes + final sweep, all findings fixed. The fake-host suite
  caught 2 real bugs, fixed in v1.0.1 (per-row borders >100 rows; orphaned on-edit
  handler). `npm run build` + both manifest validations green.
- NOT verified (needs real Excel, ~10 min when Daniel has the machine free): waterfall
  rendering + one manual "Set as Total", shortcut conflict dialogs (Ctrl+Shift+V/Z/C and
  format keys - deliberate UpSlide-style shadowing), Office.js load-ordering (fake sync
  is a no-op), copyFrom relative-ref rewrite.

## Next actions (all Daniel-gated)

1. Sideload pass: `cd ~/plsfix && npm start` - checklist in tasks/v1-plan.md
   "Remaining for launch". Report anything odd; fixes are cheap.
2. Hosting: pick suite path key (proposal /modelis/), EXCLUDE it from Cloudflare Access
   (auth walls break panes - docs/research/launch-path.md), deploy dist/ + shortcuts.json
   + icons to the suite host, swap manifest.prod.xml if the key differs.
3. M365 centralized deployment: admin center > Integrated apps > Add-ins > Upload Custom
   Apps (manifest.prod.xml), assign to a group. 24-72h propagation. JS-only updates need
   no admin action afterward.
4. git tag v1.0.1 + GitHub remote (none exists yet) when Daniel wants it.
5. Later/AppSource: needs real support + privacy pages first (research file has the
   rejection checklist). v1.x feature candidates: docs/FEATURES.md section 11 (SMT.ROUND,
   reconciliation solver, tornado, unpivot).

## Gotchas for the next session

- README.md + ROADMAP.md are Daniel's hand-curated files - surgical edits only.
- `.claude/` is gitignored and vitest include is pinned (agent worktrees nest inside the
  repo; never `git add -A` with an un-ignored worktree present).
- Excel read-back is never exact-matched without canonicalization (cycles.ts
  canonicalNumberFormat); [hidden] needs the !important reset (styles.css); full lesson
  list in tasks/lessons.md.
- manifest.xml carries duplicated V1_0 + nested V1_1 blocks - edit BOTH or they drift.
- Versioning: patch per reviewed merge, footer shows v1.0.001-style, manifest stays
  4-part (1.0.0.0).
- Builder-agent pattern that worked: opus worktree agents, one chunk each, merged
  sequentially with gates; keep-both conflict resolution can chop function tails - check
  tsc after every merge.
