# Todo

REOPENED 2026-08-28 evening for v2 (Excel -> PowerPoint tracked links). v1.1.0 hosted live
at dbautomatizacijas.com/modelis/. Open v1 gates (Daniel's): Excel visual pass + M365 admin
upload. Design: docs/superpowers/specs/2026-08-28-ppt-links-design.md. Resume from tasks/AUTORESUME.md.

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

## Hosted launch (27.08 night) - LIVE at /modelis/

- [x] Rust static host server/ (:8804) + suite registration (key modelis) + deploy.sh
      target + CF Access bypass; public edge verified, rest of suite still gated.
- [x] Daniel's wef carries manifest.prod.xml -> add-in loads from the server at every
      fresh Excel launch. v1.1.0.
- [ ] DANIEL: quit + reopen Excel -> the add-in's tab; visual pass (v1-plan checklist)
- [ ] DANIEL/IT: M365 admin upload of manifest.prod.xml (steps in v1-plan)

## Full crash debug (27.08 evening) - COMPLETE at v1.0.2

- [x] Crash located: npm start sideloaded before the dev server was up (no
      dev_server_port config) + vite bound IPv6-only on Node 25. Both fixed + verified.
- [x] Everything else ruled out with evidence: manifest/shortcuts clean, pane boots in
      a real browser, excel.ts proven free of unloaded reads (strict-load fake host +
      mutation sweep, suite 219 green). Details: tasks/v1-plan.md "Full crash debug".

Resume pin: tasks/AUTORESUME.md

## Review

(to be filled at milestone end)

## v2: Excel -> PowerPoint tracked links (28.08, approved plan)

Tier 1 groundwork (clean code the feature needs anyway)
- [x] T1.1 split src/excel.ts -> src/excel/{selection,formatting,fill,autocolor,trace,charts,workbook}.ts + index.ts barrel (no test changes)
- [x] T1.2 manifest generator: manifest/spec.ts -> scripts/build-manifests.ts -> manifest.xml (dev) + manifest.prod.xml; manifest:check
- [x] T1.3 dev loop: wef hard link -> plain copy of manifest.prod.xml; npm start links dev manifest; poststop restores prod; CLAUDE.md/README corrected
- [x] T1.4 one version source (package.json) + version:check (Cargo.toml, manifests, footer)
- [x] T1.5 npm run check (tsc, eslint strict, prettier, vitest, cargo test, manifest:check, version:check) + GitHub Actions
- [x] T1.6 src/ui/report.ts error surface (onerror + unhandledrejection, Copy details)
- [x] T1.7 src/ui/{toast,guard,tabs}.ts extracted from main.ts with tests

Tier 2 feature (v2.0.000)
- [ ] T2.0 spike in real PowerPoint/Excel (spec section 10) -> research corrections + fallback go/no-go
- [ ] T2.1 src/link/{model,crypto,status}.ts + tests
- [ ] T2.2 src/link/relay.ts RelayClient + fetch-stub tests
- [ ] T2.3 server: store.rs (sqlite) + relay.rs routes + sweeper + tests; MODELIS_DATA; systemd unit
- [ ] T2.4 test/fakeppt.ts fake PowerPoint host (strict load semantics)
- [ ] T2.5 src/ppt/host.ts adapter (scan/insert/refresh/break, <1.8 fallback) + integration tests
- [ ] T2.6 src/excel/links.ts adapter (anchors, registry, render, export/push/remove) + integration tests
- [ ] T2.7 PowerPoint pane (pptpane.html, src/ppt/main.ts: Inbox, Links, Update, Settings) + Excel Links tab
- [ ] T2.8 manifest: Presentation host + ribbon; Vite second entry; server cache prefix; deploy
- [ ] T2.9 docs: feasibility corrections, CLAUDE.md relay exception, README (surgical), AUTORESUME; tag v2.0.000
- [ ] DANIEL: real Office pass (spec section 11) after deploy

Tier 3 (after v2.0): brand settings per workbook; ribbon coverage; ROADMAP sync (FEATURES.md deltas); SMT.ROUND, reconciliation solver, tornado, unpivot; native PPT tables (kind "table")
