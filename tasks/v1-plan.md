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
- [ ] C2 Format cycles: number-format families (general/currency/percent/multiple/date),
      title/result/item row styles, fill/font/border cycles; palette-driven; cycle state
      machine in pure TS + tests; shortcuts wired.
- [ ] C3 Autocolor v2: same-sheet vs cross-sheet vs external-file vs partial-input
      classification, legend + per-type overrides in Brand tab, optional autocolor-on-edit
      (onChanged, off by default).
- [ ] C4 Auditing: formula consistency overlay (R1C1 compare, striped/solid fills,
      save/restore formats) + Smart Track pane (getDirectPrecedents/Dependents tree,
      keyboard navigation).
- [ ] C5 Fill & paste: fast-fill auto-extent, paste values/formats/transpose/skip-blanks,
      preserve-formulas paste; quick CAGR, sign flip, decimal steppers, IFERROR unwrap +
      custom fallback.
- [ ] C6 Charts: waterfall builder from a bridge table, CAGR arrow, chart brand format.
- [ ] C7 Workbook tools: TOC sheet generator, sheet explorer pane, name scrubber.
- [ ] C8 Launch polish: gap-research features triaged in, onboarding hints, docs,
      production manifest (non-localhost URLs decision), README, full QA sweep,
      reviewer sweep across the whole codebase, v1.0.0.

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
