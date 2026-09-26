# pls,fix video: the 87,5 s motion graphic

Builds `~/Desktop/pls,fix video.mp4`: 1920x1080, 60 fps, H.264, an original soundtrack (AAC
stereo, -16 LUFS), Latvian captions. The intro's "pls fix" mail, the Excel hero (Autocolor, Audit
overlay, Precedents, Model check), six toolbox chapters (formats, charts, templates, Find a
combination, Super Find, brand), the Excel to PowerPoint link (export, insert, a changed number,
Push all, Update all, the object tools), the end card with the GitHub address. Every app frame is
pls,fix running in Excel and PowerPoint for the web on the demo workbook and deck.

Rust (`plsfix-video`) captures, renders, audits and delivers; the animation is HTML/CSS/JS in
`comp/` (the browser is the renderer); the capture's browser moves are Playwright in `capture/`
(the pane lives in cross-origin frames, which the repo's web rig already reaches that way). Spec:
`docs/superpowers/specs/2026-09-26-plsfix-video-design.md`. Engine: the Planner video's
(`~/planner/video` at 2093853), recipe: the `making-app-videos` skill.

## Run

```
cd video
cargo run --release -- all          # capture + render + audit + deliver
cargo run --release -- capture      # real screens -> build/shots/ (+ shots.json)
cargo run --release -- music        # the soundtrack alone -> build/music.wav (render runs it first)
cargo run --release -- render       # build/plsfix-video.mp4 (--draft: 960x540 at 30 fps)
cargo run --release -- audit        # the gate, exit 1 on any finding
cargo run --release -- deliver      # audit, then copy to the Desktop (never over a changed copy)
cargo run --release -- all --track ~/Music/song.mp3   # any song in place of the score
cargo run --release -- still 9.9 44.2   # single frames to build/stills/ for QA
cargo test
```

Needs Google Chrome, ffmpeg + ffprobe, Montserrat, Node, and the web rig of
`scripts/rig/README.md`: the scratch profile `~/.cache/plsfix-rig-chrome` signed in to Office
once by hand and the dev certificates (`npm start` once). `capture` starts the scratch Chrome
(headless, CDP 9222) and the manifest server (3001) when they are not already up, and stops what
it started. The OneDrive site comes from `PLSFIX_RIG_SITE` or the profile's history; the demo
files go to a fresh `plsfix-video/run-<time>` folder there (a closed tab keeps its file locked
for a while), older run folders are removed. Each run generates a new link key and pairs
PowerPoint with it, so the Inbox holds only that run's export.

## Change something

- A caption: `comp/copy.lv.json` (the audit lints it: no em dash, no rejected words).
- A timing: the `K` table at the top of each scene in `comp/js/scenes/`; spans in `comp/js/timing.js`.
- A screen: `capture/excel-shots.mjs` or `capture/ppt-shots.mjs` (what to press, which
  rectangles to record), then `capture` again. A selector that stops matching fails by name.
- The music: `src/core/sound/` as in the Planner engine; `--track <file>` swaps in any song.

## The gate (`audit`)

Format (h264, 1920x1080, 60/1, yuv420p, BT.709 tags, 80-90 s, exact frame count), motion (no
still stretch over 1,5 s before the end card, no black outside the fades), colour (browser frame
and decoded video agree within 5 at four probe points), determinism (ten frames drawn in opposite
orders by two pages match up to renderer noise: at most 200 pixels off by more than 32 levels), layout (every caption visible, inside the title-safe
area, not covered), sound (-16 LUFS within 1,5, true peak at most -1 dBTP, in step with
`build/music.wav` within 2 ms at 5,0 s and 64,5 s), copy, provenance (the demo build's sha256,
Office for the web, the pane's own version).
