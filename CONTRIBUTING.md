# Contributing to pls,fix

Issues and pull requests are welcome at https://github.com/bendiksdaniels/plsfix. This page is
the short version of how the repo works; `CLAUDE.md` holds the full map (every module, the data
flow, where bugs live).

## Set up

Requirements: Node.js 22, npm, a stable Rust toolchain, and desktop Excel with Microsoft 365 for
a sideload.

```bash
npm install
npm run check        # every gate: tsc, eslint, prettier, vitest, cargo test, manifests, version
npm start            # sideloads manifest.xml into desktop Excel, served by vite on https://localhost:3000
npm run start:ppt    # the same for PowerPoint
npm stop             # ends the sideload and puts the production manifest back
```

`npm run demo` builds the demo workbook (`demo/`, Rust) and opens Excel and PowerPoint on it.
A relay for local development: `MODELIS_DATA=./data cargo run --manifest-path server/Cargo.toml`
(vite proxies `/api` to it).

## Layout

- `src/` the panes (TypeScript, Vite): `main.ts` + `pane/` the Excel pane, `ppt/` the PowerPoint
  pane, `excel/` the Office.js adapter, `link/` the pure link model and crypto, `functions/` the
  custom functions, pure modules beside them. Tests sit next to the code (`*.test.ts`) and in
  `test/` (a fake Excel and PowerPoint host with strict load semantics).
- `server/` the Rust host and relay (axum, sqlite); `demo/` and `manual/` are Rust generators.
- `manifest/spec.ts` generates BOTH manifests (`npm run manifest:build`); never edit
  `manifest.xml` or `manifest.prod.xml` by hand. `public/shortcuts.json` generates the shortcut
  card the same way.
- `deploy/` the systemd unit, the Dockerfile and the compose file; `docs/` design notes, research
  and the guides; `tasks/` the working notes.

## Rules the gates enforce

- Every Office.js scalar read is preceded by `load()` + `sync()`; the fake host throws otherwise.
- A writing Excel flow ends on `syncWrite`/`paintSync` (`src/excel/protection.ts`), never a bare
  `context.sync()`; every PowerPoint sync runs under `withSyncDeadline`.
- Files stay under 400 lines, functions under 50, one concept per file, a 2-4 line header comment
  on every source file.
- `npm run check` green before a pull request; `npm run ux:check` and `npm run ux:sweep` for
  anything that touches the panes' HTML or CSS.
- `package.json` is the one typed version; `sh scripts/release.sh patch|minor` bumps, commits and
  tags after a reviewed merge, and the release workflow publishes the tag.

## Pull requests

Small, one concern each, with the test that fails without the change. Describe the behaviour
change, not the diff. Commit subjects in the imperative.
