# pls,fix video: design

Daniel, 26.09.2026: "lets make a video for plsfix ... show all the features keep it under 90s", with
the App video playbook (`~/.claude/skills/making-app-videos/`). The request is the go (one shot).

## Outcome

One MP4, `~/Desktop/pls,fix video.mp4`: 1920x1080, 60 fps, H.264, 86,0 s (under 90 by his ask; the
playbook's 60 s cap is his to lift), an original soundtrack. Captions carry the story in Latvian; the
add-in stays English, as it is. Every app frame is a real screenshot of pls,fix v2.9.1 running in
Excel and PowerPoint for the web on a fresh upload of the demo workbook (`npm run demo:build`,
DemoCo SIA is invented) and the tracked demo deck. No personal data: the Office header (account,
file path) is cropped away, the SharePoint address never shows.

## Storyboard (seconds; 120 BPM grid, 1 beat = 0,5 s)

| Time | Scene | Motion | Caption |
|---|---|---|---|
| 0-6,4 | Intro | a mail notification drops in: "pls fix"; a mint comma falls between the words ("pls,fix"), the card melts, the comma grows into the mark (navy square), the wordmark slides out; three words; the claim; the camera dives into the mark | "Formatē. Pārbaudi. Sasaisti." then "Mazāk *pls fix* e-pastu." |
| 6-26,4 | Excel hero (P&L) | the navy lifts off a floating Excel window (the pls,fix ribbon tab and pane open); the pointer runs Autocolor (the fonts recolour, wiping across the block), Audit overlay (stripes; the planted hardcode ringed), Precedents (the trace list; the source cells selected), then the Workbook tab's Model check with a push-in on its list | "Rīki finanšu modeļiem, *tieši Excel*" . "*Autocolor*: ievades zilas, formulas melnas" . "*Audit overlay*: kļūda izceļas pati" . "*Precedents*: redzi, no kurienes nāk skaitlis" . "*Model check* atrod to, ko pamanītu MD" |
| 26-60,4 | Toolbox | the window drops into a 3x2 contact sheet of six chapters; each chapter takes the stage ~5,4 s with a chip and one caption, a real before/after inside: format cycles, the three charts, a template block, Find a combination, Super Find, the Brand tab | "Vairāk nekā *100 rīku*" . six chapter lines |
| 60-80,4 | PowerPoint | "Un jā, arī PowerPoint"; Excel and PowerPoint windows side by side; a chart is exported, flies across, is inserted on the slide; a number changes in Excel, Push all, Update all, the slide's chart follows; the Tools tab lines shapes up | "Grafiks no Excel uz slaidu" . "Slaidā rediģējams, nevis bilde" . "Maini skaitli Excel, slaids atjaunojas" . "Un rīki slaidu kārtošanai" |
| 80-86 | Outro | the mark lands, the wordmark, the claim returns, the line and the address, a slow lean-in to the curtain | "Mazāk *pls fix* e-pastu." . "Bezmaksas. Windows, Mac un tīmeklī." . "github.com/bendiksdaniels/plsfix" |

"Vairāk nekā 100" is counted, not guessed: 90 distinct `data-action` tools in `taskpane.html`, the
Links tab's export and list actions, the PowerPoint pane's object tools and link actions.

## Look

The pls,fix sheet, not Signet's (the add-in is Daniel's own product, no sponsor): ground a deep navy
field from the brand navy #14213D, mint #2EC4B6 for the mark's comma, accents and live values,
stone #F4F4F4 text, a coral ring only for the planted error. Montserrat 400/600/700 for type (the
engine checks it), the comma in Georgia bold as in `public/assets/icon.svg`. Motion as in the
Planner video: expo/quint entrances, springs, 3D-tilted windows with a slow drift, masked
word-by-word headlines, directional blur on whip moves, push-ins, lifted crops with a mint edge.

## How it is built

`video/` in this repo, the Planner's engine (`~/planner/video` at 2093853) renamed `plsfix-video`.

- `capture` drives Excel and PowerPoint for the web through the repo's own rig (`scripts/rig/`,
  the scratch Chrome profile signed in to Office, headless, on CDP 9222): it builds the demo
  workbook, uploads it and the demo deck into a `plsfix-video` OneDrive folder (overwriting only
  that folder's two files), registers the add-in from `manifest.prod.xml`, then walks the shot list
  in `video/capture/*.js`: select a range through Office.js inside the pane frame, press the real
  button, wait for the pane's toast, screenshot at 1440x948 CSS at 2x and crop the 48 px Office
  header. Rectangles come from the DOM (pane) and from `Range.left/top/width/height` plus the grid
  origin (cells). Playwright, not the engine's Rust CDP: the pane lives in cross-origin frames the
  rig already knows how to reach. The Rust command runs the snippets, checks every file and writes
  `shots.json` (the demo's sha256, the pane version read from the pane footer, the date).
- `comp/`: the composition, classic scripts, every property a pure function of t.
- `render`, `music`, `audit`, `still`, `deliver` as in the Planner engine.

## The gate (`plsfix-video audit`)

As the Planner's, with two changes that follow from this video: the length window is 80-90 s, and
provenance is the demo workbook's sha256 on the uploaded copy plus the pane's own version (the host
is Microsoft's, so "127.0.0.1" cannot hold; the add-in's data is the demo's alone).

## Soundtrack

The Planner score and sound design, unchanged: sections intro 0, groove 6, lift 26, break 60,
launch 64, outro 80; cues from each scene's timing table. -16 LUFS, AAC 192k 48 kHz. Measured, not
heard; `--track <file>` swaps in any song.

## Out of scope

An English cut (one copy file away), desktop Office captures (his Office stays hands-off), a
release asset or any upload of the video.
