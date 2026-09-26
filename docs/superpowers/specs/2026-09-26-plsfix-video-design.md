# pls,fix video: design

Daniel, 26.09.2026: "lets make a video for plsfix ... show all the features keep it under 90s", with
the App video playbook (`~/.claude/skills/making-app-videos/`). The request is the go (one shot).

## Outcome

One MP4, `~/Desktop/pls,fix video.mp4`: 1920x1080, 60 fps, H.264, 87,5 s (under 90 by his ask; the
playbook's 60 s cap is his to lift), an original soundtrack. Captions carry the story in Latvian; the
add-in stays English, as it is. Every app frame is a real screenshot of pls,fix v2.9.1 running in
Excel and PowerPoint for the web on a fresh upload of the demo workbook (`npm run demo:build`,
DemoCo SIA is invented) and the tracked demo deck. No personal data: the Office header (account,
file path) is cropped away, the SharePoint address never shows.

## Storyboard (seconds; 120 BPM grid, 1 beat = 0,5 s)

| Time | Scene | Motion | Caption |
|---|---|---|---|
| 0-6,4 | Intro | a mail notification drops in: "pls fix"; a mint comma falls between the words ("pls,fix"), the card melts, the words fly into the wordmark as the mark (navy square) pops in; three words; the claim; the camera dives into the mark | "Formatē. Pārbaudi. Sasaisti." then "Mazāk *pls fix* e-pastu." |
| 6-26,4 | Excel hero (P&L) | the navy lifts off a floating Excel window (the pls,fix ribbon tab and pane open); the pointer runs Autocolor (the fonts recolour, wiping across the block), Audit overlay (stripes; the planted hardcode ringed), Precedents (the trace list; the source cells ringed, arrows drawn), then the Workbook tab: Run model check is pressed, the pane scrolls to the findings and they lift row by row under a push-in | "Rīki finanšu modeļiem, *pašā Excel*" . "*Autocolor*: ievaddati zili, saites zaļas, formulas melnas" . "*Audit overlay*: kļūda izceļas pati" . "*Precedents*: redzi, no kurienes nāk skaitlis" . "*Model check* atrod to, ko pamanītu MD" |
| 26-60,4 | Toolbox | the window drops into a 3x2 contact sheet of six chapters; each takes the stage 5,3 s with a chip and one caption, a real before/after inside: the Title preset, a waterfall from the bridge table (the demo charts step aside), a DCF block on a fresh sheet, Find a combination (the lines and the Target ringed), Super Find for "EBITDA 2024" (hits on the bridge table and on the hidden Scratch sheet), the Brand tab (colours and number preview ringed) and its keyboard shortcuts | "Vairāk nekā *100 rīku*" . Formāti "Firmas stils ar vienu klikšķi" . Grafiki "*Waterfall* no tabulas, firmas krāsās" . Šabloni "Gatavs *DCF* bloks ar dzīvām formulām" . Saskaņošana "Atrod, kuri skaitļi kopā dod summu" . Viss fails "Atrod jebko visās lapās, arī slēptajās" . Zīmols "Tavas krāsas, tava valoda, tavi taustiņi" |
| 60-80,4 | PowerPoint | "Un jā, arī PowerPoint"; Excel and PowerPoint side by side; a chart is exported, flies across, is inserted on slide 3; revenue growth goes from 8 % to 25 % in Excel, Push all on the Links tab; PowerPoint's Links tab reads "Update available", Update all, the bars rise while the camera pushes in on the slide; the Tools tab aligns and spaces three shapes | "Grafiks no Excel uz slaidu" . "Ievietots slaidā izvēlētajā vietā" . "Maini skaitli Excel, slaids atjaunojas" . "Un rīki objektu līdzināšanai slaidā" |
| 80-87,5 | Outro | the mark pops in as the PowerPoint act leaves, the wordmark slides out, the claim returns, then free and open source with the GitHub address (Daniel, mid-build: "mention that it open source on github"), held about 4 s, a slow lean-in, the curtain over the last 0,3 s | "Mazāk *pls fix* e-pastu." . "Bez maksas, ar atvērtu kodu" . [GitHub] "github.com/bendiksdaniels/plsfix" |

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
  workbook, uploads it and the demo deck into a fresh `plsfix-video/run-<time>` OneDrive folder (older
  run folders removed; SharePoint keeps a closed tab's file locked for a while, so no overwriting
  in place), registers the add-in from `manifest.prod.xml`, generates a fresh link key and
  pairs PowerPoint with it (so the Inbox holds only this run's export), then walks the shot list
  in `video/capture/*.mjs`: select a range through Office.js inside the pane frame, press the real
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

## Review (26.09.2026)

One critic pass (one opus agent) on the first render. Fixed: Push all and Model check were
pressed on the wrong screen (new "before" shots, the real buttons pressed), the slide's update
was too small to see (the edit is now 8 % to 25 %, a push-in on the slide chart, the bars rise),
PowerPoint's Links tab opened by itself (now clicked, "Update available" ringed), Super Find did
not show a hidden sheet ("EBITDA 2024" does), rings on section titles (now on the rows), the
waterfall upstaged by the demo's Revenue chart (moved aside), the intro card flashing white, the
tagline overlapping the claim, a too-short end card (+1,5 s), empty ground beside push-ins
(leans now keep the frame filled), four captions. The determinism gate now tells renderer noise
(at most 18 levels, measured) from a leaked state (more than 32).
