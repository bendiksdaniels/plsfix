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
- `manifest.xml` = dev (localhost:3000), `manifest.prod.xml` = production URLs; both carry
  duplicated VersionOverrides V1_0 + nested V1_1 blocks - edit BOTH or they drift.
- `README.md` and `ROADMAP.md` are Daniel's hand-curated files: surgical edits only.
- Versioning: package.json + server crate stay in lockstep; pane footer shows
  vMAJOR.MINOR.PATCH with a 3-digit patch.

State pins: `tasks/AUTORESUME.md` (resume here), `tasks/v1-plan.md`, `tasks/lessons.md`.
