# Web verification rig

Drives Excel and PowerPoint for the web in a scratch Chrome over CDP, with the
add-in registered from this repo's manifest, so a link flow is proven end to end
without desktop control. These are the scripts that used to live in session
scratchpads (tasks/lessons.md, 29.08 and 30.08). Manual runs only, never CI:
it needs a signed-in browser and a live document.

## One-time setup

1. Dev certificates: `npm start` installs them once (`~/.office-addin-dev-certs`).
2. The certificate's SPKI, for Chrome:
   `openssl x509 -in ~/.office-addin-dev-certs/localhost.crt -pubkey -noout | openssl pkey -pubin -outform der | openssl dgst -sha256 -binary | base64`
3. Scratch Chrome with CDP (its own profile, so your daily browser stays out of it):
   `"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --remote-debugging-port=9222 --user-data-dir="$HOME/.cache/plsfix-rig-chrome" --ignore-certificate-errors-spki-list=<spki> --no-first-run &`
4. Sign in to Office on the web in that Chrome yourself. The rig never sees
   credentials. Open the demo workbook (`npm run demo:build`, then upload
   `demo/out/pls,fix Demo Model.xlsx` to OneDrive) and a deck.
5. `npm run rig:manifests` in a second terminal: serves `manifest.prod.xml`
   on `https://127.0.0.1:3001` with CORS, which is what the registration needs.
   Never pipe it through `head` or another short-lived reader: it logs every
   request, and the first write after the reader closes kills it (08.09).
6. Register the add-in: append to the document URL
   `&wdaddindevserverport=3001&wdaddinmanifestfile=manifest.prod.xml&wdaddinmanifestguid=FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E`
   and run `HOST_KIND=x npm run rig -- @scripts/rig/snippets/sideload.js`
   (`x` = Excel, `p` = PowerPoint). It grants local-network access over CDP,
   clicks the two registration dialogs and reports the ribbon tabs; "pls,fix"
   must be among them. The parameters stay in the URL, so a reload re-registers
   without the dialogs.

## Running

- `npm run rig -- @scripts/rig/snippets/pane-open.js`: opens the Excel pane
  through the ribbon when it is closed and returns its frame box
  (`SHOTS=1` also saves one screenshot per tab).
- `npm run rig -- '<async (page, ctx, pages, shot, h) => { ... }>'`: an ad-hoc
  snippet. `h` is `helpers.mjs` (`hostPage`, `frameByUrl`, `paneText`,
  `clickIn`, `ribbonTabs`, `shotsDir`); `shot(name)` saves a PNG under
  `.superpowers/rig/` (git-ignored).
- `snippets/proof-charts-excel.js` and `proof-charts-ppt.js`: the 30.08 chart
  proof (three exports, three inserts, a source change, Push all, Update all,
  read-back of groups against pictures). Read them before writing a new one:
  every pane action is verified from inside the pane frame with `Excel.run` or
  `PowerPoint.run`, not from screenshots.

## Rules the web taught us

- The "Enable Developer Mode" dialog carries an opt-in checkbox (Excel
  `#WACDialogOptInCheckbox-input`, PowerPoint `#optInCheckbox`); OK without
  the tick only dismisses it. The tick counts on the NEXT load of the same
  URL: reload and run `sideload.js` again until the pls,fix tab shows.
- The pane's tab strip can sit outside the frame's viewport in a narrow
  window: click pane buttons with a DOM click (`el.click()` in
  `frame.evaluate`), not Playwright's actionability click, and widen the
  window over CDP (`Browser.setWindowBounds`) so the ribbon is not
  collapsed behind its chevron when reading the tabs.
- A pane toast can be the previous action's: wait on the busy flag AND a
  toast text that differs from the one read before the click.
- Panes on the same site share localStorage, but the PowerPoint pane still
  reports "unpaired" until the key is saved in its Settings (or it reloads).
- The OneDrive "Create or upload" menu stays open across snippets: press
  Escape first.

- Keep the host tab in FRONT while its pane is busy: a `PowerPoint.run` in a
  background tab never returns. Switch tabs only when the pane is idle.
- Wait on the pane's own busy flag (`#tab-inbox` disabled) plus a toast change,
  never on a fixed delay.
- A batch that never returns jams every later write while reads still answer.
  Heal by reloading the pane frame (`location.reload()` inside it), or the deck
  with the `wdaddin*` parameters kept (`proof-charts-ppt.js` does this).
- Pre-existing cross-origin frames report an empty `url()` on a fresh CDP
  connection: match frames by `location.href` (`h.frameByUrl`).
- Charts on the web are canvas-drawn and `activate()` does not select them:
  chart export goes through the pane's chart picker (`#export-chart-pick`).
- The custom-functions runtime may not start on a web sideload, so `=PLSFIX.*`
  shows `#NAME?` there. That is the sideload, not the pane.
- Office Online sometimes answers a navigation with "services aren't available
  right now". Retry.
- Costs measured on the web (30.08): a 21-shape pie 18 s, a 20-shape column
  chart 24 s, Update all over three chart links 54 s, a text box about 1 s.

`playwright-core` is not a project dependency: `drive.mjs` finds it through the
same resolver as the UX gate (`scripts/ux/browser.mjs`).
