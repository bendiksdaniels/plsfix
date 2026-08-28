# Model Tools

Office.js Excel task-pane add-in for financial modelling (TypeScript/Vite/Vitest) plus a
tiny Rust static host (`server/`, axum, port 8804) that serves the built pane to the
DB Automatizācijas suite at `dbautomatizacijas.com/modelis/` (key `modelis`).

- Pane code in `src/`, tests in `src/` + `test/` (fake Office.js host with strict load
  semantics). `npm test`, `npm run build`, `npm run validate`. Dev sideload: `npm start`.
- `server/`: `cargo test --manifest-path server/Cargo.toml`. Serves `dist/` (env
  `MODELIS_STATIC`), `/healthz`, `/version`; gateway strips the `/modelis/` prefix.
  The public path carries a Cloudflare Access bypass - Office webviews cannot pass
  Access, so nothing private may ever be served here.
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

Entry points: `taskpane.html` + `src/main.ts` (Excel pane, web edge), `pptpane.html` + `src/ppt/main.ts` (PowerPoint pane), `server/src/main.rs` (axum host + relay), `manifest/spec.ts` -> `scripts/build-manifests.ts` (both manifests).

Modules (what each owns):
- `src/main.ts`: Excel pane wiring: `dispatch(action)` switch, ribbon/shortcut command table, brand dashboard, sheet explorer, boot in `Office.onReady`. Web edge, no business logic.
- `src/ui/`: shared pane pieces without Office.js: toast, guard, tabs, report (error surface), version.
- `src/excel/`: the Excel Office.js adapter, one file per feature area behind an `index.ts` barrel that keeps the import path `./excel`: `shared.ts` (public types + `parseAddress`), `internal.ts` (private range/fill/host-capability helpers, not barrel-exported), `undo.ts`, `selection.ts`, `formulas.ts`, `paste.ts`, `autocolor.ts`, `audit.ts`, `trace.ts`, `charts.ts`, `workbook.ts`, `links.ts` (tracked-link flows) + `link-anchors.ts` (registry, anchors, render; not barrel-exported).
- Pure model modules (typed in/out, no I/O): `settings` (brand palette), `cycles`, `paste` (fill maths), `classify`, `audit`, `chartmath`, `model`, `workbook`.
- `src/link/` (core, pure): `model` (types + codecs), `crypto` (HKDF/AES-GCM, base64url, tokens), `status` (link state rules, slide fitting), `png` (IHDR size), `workspace` (pairing key + host key store), `relay` (fetch client, typed errors).
- `src/ppt/`: `host.ts` (the only PowerPoint Office.js code: scan by tags, insert, refresh in place or by reinsertion, break, go to slide), `links.ts` (orchestration: list, update, inbox), `main.ts` (pane boot, state and one guarded handler per button), `views.ts` (DOM renderers: link table, inbox list, status and age labels), `actions.ts` (pure: row keys, which rows a button acts on, update detail lines).
- `src/pane/links-tab.ts` (planned): the Excel "Links" tab.
- `server/`: `lib.rs` router + cache middleware, `store.rs` sqlite (WAL, auth hash, 2 revs, TTL sweep), `relay.rs` `/api` routes (bearer, ETag, body limits), `main.rs` env + bind + hourly sweeper.
- `test/`: `fakehost.ts` (in-memory Excel host, strict load semantics), `fakeppt/` (PowerPoint fake, planned), `fakerelay.ts`, `fakepng.ts` (planned), `*.integration.test.ts`.

Data flow, link export (v2): selection -> `src/excel/links.ts` (hidden name anchor + registry in `workbook.settings`) -> `Range.getImage` -> `src/link/model` encodePayload -> `src/link/crypto` seal -> `src/link/relay` putLink -> `server/relay.rs` -> `store.rs` (ciphertext only) -> inbox item sealed with the workspace key. PowerPoint: `src/ppt/links.ts` listInbox -> open -> `src/ppt/host.ts` insertLink (rectangle + picture fill + tags). Update: scanLinks (tags) -> relay status -> getLink -> open -> refreshLink (`fill.setImage`, geometry untouched).

Invariants: link identity is the shape tags (`SMT_LINK`, `SMT_KEY`) in PowerPoint and the hidden name `SMT_LINK_<id8>` in Excel, never an address or a position; the server holds ciphertext and `sha256(authKey)` only; refresh never writes geometry except height on an aspect change; every Office.js scalar read is preceded by `load()` + `sync()` (the strict fakes throw otherwise); manifests are generated, never hand-edited; `package.json` is the only typed version; hand-curated files (`README.md`, `ROADMAP.md`) get surgical edits only.

Run / test: `npm start` (Excel sideload with the dev manifest), `npm run start:ppt`, `npm stop` (restores the prod manifest into wef), `npm run check` (tsc, eslint, prettier, vitest, cargo test, manifest:check, version:check), `MODELIS_DATA=./data cargo run --manifest-path server/Cargo.toml`, deploy only via the gateway (`db deploy modelis`).

Where bugs live (symptom -> file): wrong cell/format written -> `src/excel/<feature>.ts`; `PropertyNotLoaded` -> a missing `load()` in that adapter (reproduce in the fake host first); dead pane button -> `src/main.ts` dispatch or the `[hidden]` cascade in `styles.css`; link not found in PowerPoint -> tags, `src/ppt/host.ts` scanLinks; "source missing" -> anchor resolution in `src/excel/links.ts`; relay 403/404 -> key derivation `src/link/crypto.ts` or TTL in `server/store.rs`; wrong picture size -> `src/link/status.ts`; add-in not loading -> `manifest/spec.ts` + `npm run validate`; Excel crash on `npm start` -> `package.json` `config.dev_server_port` + vite IPv4 (lessons).
