# Model Tools

Office.js add-in for financial modelling with two hosts: the Excel task pane
(`taskpane.html`, model tools + Links tab) and the PowerPoint pane (`pptpane.html`, tracked
Excel->PowerPoint links), TypeScript/Vite/Vitest, plus a Rust axum host (`server/`, port
8804) that serves both panes AND the end-to-end encrypted link relay (`/api/links`,
`/api/inbox`, sqlite at `MODELIS_DATA`) to the DB Automatizācijas suite at
`dbautomatizacijas.com/modelis/` (key `modelis`).

- Pane code in `src/`, tests in `src/` + `test/` (fake Office.js host with strict load
  semantics). `npm run check` (all gates), `npm run build`, `npm run validate`. Dev sideload:
  `npm start` (Excel) / `npm run start:ppt` (PowerPoint); `npm stop` / `npm run stop:ppt`.
- `server/`: `cargo test --manifest-path server/Cargo.toml`. Serves `dist/` (env
  `MODELIS_STATIC`), `/healthz`, `/version`; gateway strips the `/modelis/` prefix.
  The public path carries a Cloudflare Access bypass - Office webviews cannot pass
  Access, so nothing readable may ever be served here. The one exception, by design:
  relay blobs under `/api/` are encrypted in the pane (AES-GCM, keys derived from a
  per-link token that exists only inside a deck's tags and the workbook registry); the
  server stores ciphertext and `sha256(authKey)` only, never plaintext or file names.
- Deploys go through the gateway's audited flow (`db deploy modelis`), which builds the
  pane locally (`npm run build`) and the server binary on the VPS. Never hand-copy files
  to production.
- `manifest.xml` (dev, localhost:3000) and `manifest.prod.xml` (prod, server URLs) are
  both generated from `manifest/spec.ts` via `npm run manifest:build` - never edit
  either file by hand. `npm stop` restores the prod manifest into the wef folders
  (`scripts/wef-restore-prod.sh`).
- `README.md` and `ROADMAP.md` are Daniel's hand-curated files: surgical edits only.
- Toolchain: `typescript` in package.json is aliased to `@typescript/typescript6` (the TS team's
  side-by-side shim) because typescript-eslint refuses TS 7; the real TS 7 compiler is
  `@typescript/native` and still provides `tsc`. Agent worktrees nest under `.claude/`, so lint
  ignores use `**/dist/**`-style globs plus `.claude/**` and `.superpowers/**`.
- Versioning: package.json + server crate stay in lockstep; pane footer shows
  vMAJOR.MINOR.PATCH with a 3-digit patch.

State pins: `tasks/AUTORESUME.md` (resume here), `tasks/v1-plan.md`, `tasks/lessons.md`.

## Map

Entry points: `taskpane.html` + `src/main.ts` (Excel pane, web edge), `pptpane.html` + `src/ppt/main.ts` (PowerPoint pane), `server/src/main.rs` (axum host + relay), `manifest/spec.ts` -> `scripts/build-manifests.ts` (both manifests), `src/functions/index.ts` -> `dist/functions.js` (Excel custom functions, second vite pass).

Modules (what each owns):
- `src/main.ts`: Excel pane wiring: `dispatch(action)` switch, ribbon/shortcut command table, brand dashboard (persisted to localStorage and to the workbook; boot adopts the workbook palette over the machine one), sheet explorer, boot in `Office.onReady`. Web edge, no business logic.
- `src/ui/`: shared pane pieces without Office.js: toast, guard, tabs, report (error surface), version, time, clipboard.
- `src/excel/`: the Excel Office.js adapter, one file per feature area behind an `index.ts` barrel that keeps the import path `./excel`: `shared.ts` (public types + `parseAddress`), `internal.ts` (private range/fill/chart-shell/host-capability helpers, not barrel-exported), `undo.ts`, `selection.ts`, `sizes.ts` (row-height and column-width cycles, outside SMT Undo: sizes are sheet state), `formulas.ts`, `paste.ts`, `autocolor.ts`, `audit.ts`, `trace.ts`, `charts.ts`, `tornado.ts` (sensitivity chart from a helper block beside the selection), `reshape.ts` (unpivot onto a new sheet), `workbook.ts`, `brand-store.ts` (brand palette JSON in `workbook.settings["smt.brand.v1"]`; the workbook wins over localStorage, which is the default for new workbooks), `links.ts` (tracked-link flows) + `link-anchors.ts` (registry, anchors, render) + `link-record.ts` (relay round trip: push, announce, publish/rollback); neither barrel-exported; also `paintbrush.ts`, `setCellProperties`; also `find.ts`.
- Pure model modules (typed in/out, no I/O): `settings` (brand palette), `cycles`, `paste` (fill maths, number-format decimals, formula strings), `classify`, `audit`, `chartmath` (bridge, CAGR, tornado ranking), `model`, `reshape` (wide to long), `workbook`, `rounding` (largest-remainder allocation: the parts add up to the rounded total); also `paintbrush`, `smt.paint.v1`; also `find`.
- `src/functions/`: the Excel custom functions `=SMT.ROUND` / `=SMT.ROUNDSUM`, the only code Office loads into the custom-functions runtime. `index.ts` associates them over `src/rounding.ts` (no DOM, no Office.js, no pane imports) and is built to a standalone `dist/functions.js` by `vite.functions.config.ts`; `metadata.ts` is the typed source of `public/functions.json`, written by `scripts/build-functions-metadata.ts` and gated by `npm run functions:check`. The manifest wires both files plus the `SMT` namespace into the Workbook host (`<AllFormFactors>`), reusing the pane's shared runtime.
- `src/link/` (core, pure): `model` (types + codecs), `crypto` (HKDF/AES-GCM, base64url, tokens), `status` (link state rules, slide fitting), `png` (IHDR size), `workspace` (pairing key + host key store), `relay` (fetch client, typed errors).
- `src/ppt/`: `host.ts` (the only PowerPoint Office.js code: scan by tags, insert, refresh in place - one shape or a whole update-all batch in a single sync - or by reinsertion, break, go to slide, read the active slide), `shapes.ts` (where a shape sits: the shape property list, the PowerPointApi check, walking into groups and addressing a shape through its `groupPath`), `links.ts` (orchestration: list, update, inbox), `main.ts` (pane boot, state and one guarded handler per button), `views.ts` (DOM renderers: link table, inbox list, status and age labels), `actions.ts` (pure: row keys, which rows a button acts on, update detail lines).
- `src/pane/links-tab.ts`: the Excel "Links" tab (export, workbook link list, push, go to source, remove, link key); `installLinksTab` is the only thing `main.ts` adds for it.
- `server/`: `lib.rs` router + cache middleware, `store.rs` sqlite (WAL, auth hash, 2 revs, TTL sweep), `relay.rs` `/api` routes (bearer, ETag, body limits), `main.rs` env + bind + hourly sweeper.
- `test/`: `fakehost.ts` (in-memory Excel host, strict load semantics), `fakeppt/` (PowerPoint fake) + `fakerelay.ts` + `fakepng.ts`, `ppt.support.ts` (shared `bootPpt`/`seedLink`/`pushAgain`/`memoryStore` for the `ppt.*.integration.test.ts` suites, split by feature so each file stays under 400 lines; `ppt.perf.integration.test.ts` holds the round-trip budgets, counted by `helpers.syncCount()`), `*.integration.test.ts`.

Data flow, link export (v2): selection -> `src/excel/links.ts` (hidden name anchor + registry in `workbook.settings`) -> `Range.getImage` -> `src/link/model` encodePayload -> `src/link/crypto` seal -> `src/link/relay` putLink -> `server/relay.rs` -> `store.rs` (ciphertext only) -> inbox item sealed with the workspace key. PowerPoint: `src/ppt/links.ts` listInbox -> open -> `src/ppt/host.ts` insertLink (rectangle + picture fill + tags). Update: scanLinks (tags, groups walked three levels deep on PowerPointApi 1.8) -> relay status -> getLink -> open -> `refreshLinks` (`fill.setImage` on the shape each `groupPath` leads to, geometry untouched, the whole batch in one `PowerPoint.run` and one sync; a host below 1.8, or one that refuses a shape, falls back to `refreshLink` per row).

Invariants: link identity is the shape tags (`SMT_LINK`, `SMT_KEY`) in PowerPoint and the hidden name `SMT_LINK_<id8>` in Excel, never an address or a position; the server holds ciphertext and `sha256(authKey)` only; refresh never writes geometry except height on an aspect change; every Office.js scalar read is preceded by `load()` + `sync()` (the strict fakes throw otherwise); manifests are generated, never hand-edited; `package.json` is the only typed version; hand-curated files (`README.md`, `ROADMAP.md`) get surgical edits only.

Run / test: `npm start` (Excel sideload with the dev manifest), `npm run start:ppt`, `npm stop` (restores the prod manifest into wef), `npm run check` (tsc, eslint, prettier, vitest, cargo test, manifest:check, version:check), `MODELIS_DATA=./data cargo run --manifest-path server/Cargo.toml`, deploy only via the gateway (`db deploy modelis`).

Where bugs live (symptom -> file): wrong cell/format written -> `src/excel/<feature>.ts`; `PropertyNotLoaded` -> a missing `load()` in that adapter (reproduce in the fake host first); dead pane button -> `src/main.ts` dispatch or the `[hidden]` cascade in `styles.css`; link not found in PowerPoint -> tags, `src/ppt/host.ts` scanLinks; "source missing" -> anchor resolution in `src/excel/links.ts`; relay 403/404 -> key derivation `src/link/crypto.ts` or TTL in `server/store.rs`; wrong picture size -> `src/link/status.ts`; add-in not loading -> `manifest/spec.ts` + `npm run validate`; Excel crash on `npm start` -> `package.json` `config.dev_server_port` + vite IPv4 (lessons).
