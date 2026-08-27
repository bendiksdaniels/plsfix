# Todo

## Rename + relocate (2026-08-27)

- [x] Move to `~/plsfix`, Desktop symlink kept
- [x] Rename ModelCraft -> Model Tools (package.json, package-lock.json, manifest.xml, README.md, taskpane.html)
- [x] Manifest validation green: Version 1.0.0.0 (store format; product version stays in package.json), PNG icon
- [x] Build green: src/vite-env.d.ts for vite/client types
- [x] git init + initial commit

## Brand dashboard (2026-08-27)

- [x] Tabbed task pane (Tools | Brand), charcoal/bronze reskin
- [x] Palette editor: 5 semantic slots, pickers + hex, logo upload -> local dominant-color extraction
- [x] Font + currency settings; presets/autocolor/currency format read the palette (src/settings.ts, TDD, 17 tests)
- [x] localStorage persistence + JSON import/export, live ledger preview
- [x] Code review pass: 2 real findings (hidden-vs-display cascade broke tab switching; stale logo swatches after reset) — fixed in v0.2.1, lesson captured in tasks/lessons.md. In-Excel visual check pending next sideload.

## Competitor feature research

- [x] Macabacus deep inventory (agent, 173k tok) -> docs/research/macabacus-upslide-inventory.md
- [x] UpSlide help-center full sweep (agent, 111k tok) -> docs/research/upslide-helpcenter-inventory.md
- [x] Office.js feasibility study (agent, 235k tok) -> docs/research/officejs-feasibility.md
- [x] Sales-team training files extracted -> docs/research/upslide-*-training-dump.md
- [x] docs/FEATURES.md consolidated matrix + ranked v0.3 shortlist; spot-checked vs live sources
- [ ] Roadmap deltas: proposed in FEATURES.md, awaiting Daniel's sign-off (ROADMAP.md is his file)
- [ ] Known conflict: UpSlide training file vs help center swap Autocolor/Smart Track shortcuts (remappable; flagged in research notes)

## v1 campaign (27.08, loop) - COMPLETE at v1.0.1

- [x] C1-C8 built, merged, gated (tasks/v1-plan.md); final sweep findings applied -> v1.0.0-rc
- [x] Reviews: 5 passes + final sweep; findings fixed (2 refuted with evidence)
- [x] Simulated host: test/fakehost.ts + 110 end-to-end tests (212 total green); caught 2
      real bugs, fixed in v1.0.1
- [x] manifest.prod.xml valid (path /modelis/ = proposal), README launch sections
- [ ] DANIEL: sideload pass (npm start) - checklist in v1-plan "Remaining for launch"
- [ ] DANIEL: suite-path key + Cloudflare Access exclusion -> deploy -> M365 centralized deployment

Resume pin: tasks/AUTORESUME.md

## Review

(to be filled at milestone end)
