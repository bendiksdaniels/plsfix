# AUTORESUME - Model Tools (pinned 2026-08-27 night, post-hosting)

## State: v1.1.0, HOSTED LIVE at dbautomatizacijas.com/modelis/, M365 upload = last gate

- HOSTING LIVE (Daniel: "make it easier to show up"): Rust static host `server/`
  (axum :8804) serves dist/ on the suite as key `modelis` (hidden, no sidebar
  injection); Cloudflare Access BYPASS app for the path (Office webviews cannot pass
  Access) - verified publicly, /excel/ still gated. Deploy: `deploy.sh modelis`
  (builds pane on the Mac, rebuilds host on the VPS). Daniel's wef folder now carries
  manifest.PROD: every fresh Excel launch loads the add-in from the server, no local
  servers. CAUTION: `npm stop` deletes the wef entry - re-copy manifest.prod.xml
  after dev sessions. Remaining: M365 admin upload (steps in tasks/v1-plan.md) +
  Daniel's visual pass after a full Excel restart.

- Repo `~/plsfix` (Desktop symlink), git local-only, tree clean.
- FULL CRASH DEBUG done (Daniel: "find where it crashes"): the crash was npm start
  sideloading Excel before the dev server was up (no dev_server_port config ->
  office-addin-debugging never waits) plus vite binding IPv6-only on Node 25. Both
  fixed (package.json config block; dns ipv4first in vite.config.ts) and verified
  headlessly with `npm start -- --no-sideload`. Everything else ruled out: manifest +
  shortcuts clean, pane boots in a real browser, and src/excel.ts is PROVEN free of
  unloaded reads (strict-load fake host + 60-load mutation sweep). Details in
  tasks/v1-plan.md "Full crash debug".
- All 8 v1 chunks shipped (see tasks/v1-plan.md for per-chunk detail): shared runtime +
  ribbon + 34 shortcuts, format cycles, autocolor v2 + color key, audit overlay + Smart
  Track, SMT Undo + paste suite + fast fill, native waterfall + chart tools, Workbook tab
  (TOC, sheet explorer, name scrubber), launch polish.
- Verification: 219 vitest green = 102 pure-logic + 110 end-to-end + 7 instrument
  tests. The integration suite now runs under STRICT LOAD SEMANTICS
  (test/fakehost.ts enableStrictLoadSemantics(): scalar proxy reads throw
  PropertyNotLoaded unless loaded AND synced, like real Excel) - load-ordering is no
  longer an unverified class. Earlier: 5 review passes + final sweep; fake host caught
  2 real bugs fixed in v1.0.1. `npm run build` + both manifest validations green.
- NOT verified (needs real Excel, ~10 min when Daniel has the machine free): waterfall
  rendering + one manual "Set as Total", shortcut conflict dialogs (Ctrl+Shift+V/Z/C and
  format keys - deliberate UpSlide-style shadowing), write payload size on huge
  selections, multi-area selections, copyFrom relative-ref rewrite.

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
