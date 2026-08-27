# v1 plan (loop campaign, started 2026-08-27)

Goal: launchable v1.0.0 of the Excel modelling core. Daniel's directives: loop until done,
use as many agents as needed, fold in competitor/user-gap research, keep it clean.

Quality bar per chunk: TDD for pure logic, gates green (vitest + tsc/vite build + manifest
validate), reviewer-agent pass on the diff, terse commit, version bump per Daniel's scheme
(patch per reviewed merge, minor per chunk). UI chunks additionally need a rendered check
(browser dev mode or sideload) before their roadmap box is ticked — see tasks/lessons.md.

## Chunks

- [x] C1 Foundation: shared runtime, the add-in's ribbon tab (function commands), keyboard
      shortcuts via ExtendedOverrides (v1.1 nested overrides). v0.3.0.
      In-Excel verification of shortcuts/ribbon PENDING next sideload.
- [x] C2 Format cycles: number-format families (general/currency/percent/multiple/date),
      title/result/item row styles, fill/font cycles; palette-driven; pure cycles.ts + 17
      tests; ten shortcuts wired. Merged 5d786e8, v0.4.0. Reviewer pass + in-Excel check
      pending. Open: standalone border cycle (borders ship inside row styles for now).
- [x] C3 Autocolor v2: classify.ts (13 tests), external/partial palette slots, on-edit
      toggle, insert-color-key block. Merged bea665c, v0.5.0. Reviewer + in-Excel check
      pending. Open follow-ups: structured table refs (Table1[Col]) classify as external;
      color key doesn't autosize columns (deliberate).
- [x] C4 Auditing: audit.ts (11 tests) pattern-fill overlay with snapshot/restore,
      Smart Track pane w/ chips + back stack, runtime requirement detection (1.12/1.13)
      without raising the manifest floor. Merged 4abe5a8, v0.6.0. Reviewer + in-Excel
      check pending. snapshotFills/restoreFills ready to generalize into C5's SMT Undo.
- [x] C5 Fill & paste: paste.ts (17 tests), fast-fill auto-extent, copy/paste suite incl.
      exact-formulas, CAGR/sign/decimal steppers, IFERROR toggle, one-slot SMT Undo wired
      into all 17 mutating actions. Merged 80dcf81, v0.7.0. Reviewer + in-Excel check
      pending. Known limits for C8: undo skips borders/alignment; shortcut conflicts with
      native Ctrl+Shift+V/Z/C/3 need the sideload conflict-dialog check.
- [x] C6 Charts: chartmath.ts (12 tests), NATIVE waterfall from a bridge table (typings
      rule out array-backed series) with per-point brand colors + delta reconciliation
      toast, brand chart formatter, CAGR label shape. Merged 5dbb0da, v0.8.0. Known
      defect for C8/README: Office.js has no "set as total" flag, so the closing total
      renders as a floating delta until the user right-clicks Set as Total once.
- [x] C7 Workbook tools: workbook.ts (10 tests), TOC sheet w/ marker guard + hyperlinks,
      third Workbook tab (sheet explorer w/ visibility toggles, two-step broken-name
      delete). Merged 0e9cf6c, v0.9.0 (em-dash marker fixed to hyphen at merge; two
      keep-both junction chops repaired; 102 tests green). In-Excel check pending.
- [x] C8 Launch polish -> v1.0.0-rc (559fa1a): final whole-codebase sweep (3 critical,
      4 important, 5 minor - ALL applied: audit-overlay snapshot persisted in workbook
      settings + startup restore, selection caps on reads AND writes w/ debounce,
      full-fidelity undo via setCellProperties, vite relative base for sub-path
      hosting, phantom edit-handler fix, full-combo kbd hints, prod manifest
      AppDomain/SupportUrl, dead code removed). README: shortcut table, limits,
      deploy, Mac positioning. manifest.prod.xml valid. v1.x candidates (SMT.ROUND,
      reconciliation solver, tornado, unpivot) stay in FEATURES.md section 11.

## Simulated host verification (27.08, post-rc)

- [x] test/fakehost.ts: in-memory Office.js host (2,267 lines) with host-quirk switches
      (currency-format rewrite, unfilled-fill reporting, sync-committed events) +
      test/host.integration.test.ts: 110 end-to-end tests over all 38 excel.ts exports.
      Suite total 212 green. It caught TWO real bugs, both fixed in v1.0.1 (50f0afa):
      per-row borders silently degraded above 100 rows (now per-row to 500, refused
      beyond), and a failed on-edit removal sync orphaned the handler (handle now kept
      until the removal syncs). Known simulator limits: no-op sync (load ordering not
      exercised), no formula recalculation, copyFrom does not rewrite relative refs.
      Real-Excel-only residue: waterfall rendering, shortcut conflict dialogs.

## Remaining for launch (Daniel's gates)

- [ ] Sideload pass on desktop Excel (npm start): ribbon, shortcut conflict dialogs,
      cycles, waterfall + Set as Total, undo, TOC, overlay persistence across reopen.
- [ ] Suite path key decision (proposal: /modelis/) + Cloudflare Access EXCLUSION for
      that path; deploy dist/ + shortcuts.json + icons; then M365 centralized
      deployment (admin upload manifest.prod.xml, group assignment).
- [ ] Tag v1.0.0 after the sideload pass is clean.

## Agent strategy

Research: sonnet agents (competitor deltas + user-gap mining are running). Build: C2/C3/C5/C7
are separable; dispatch worktree-isolated builder agents (sonnet/opus) once C1 foundations are
merged, one chunk per agent, merged sequentially with gates + reviewer after each merge.
C4/C6 are trickier (overlay state, chart APIs): build in main context or opus agent with review.
Fable (main context) does architecture, integration, merges, reviews, and the md files.

## Launch checklist (v1 gate)

- All chunk boxes above ticked, gates green, zero grep hits for TODO/FIXME in src.
- Sideload-verified on desktop Excel (Daniel) incl. shortcuts and ribbon.
- Production hosting decision for the manifest URLs (localhost is dev-only) + deploy story.
- README quickstart + docs/FEATURES.md updated to reflect shipped state.
- tasks/lessons.md reviewed; auto-memory updated.
