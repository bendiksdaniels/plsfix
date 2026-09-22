# Click-through of every control, 22-23.09.2026

Daniel's ask: the add-in must not crash, every button must work as intended, click through each
one, and each must be simple to use. This file is the record: what was pressed, how, what it found,
what shipped. Ledger entry: `tasks/AUTORESUME.md` (23.09). Plan: `~/.claude/plans/cheeky-orbiting-frog.md`.

## Outcome

- No crash and no dead control across both panes, both ribbons and the keyboard shortcuts, over the
  fake hosts (vitest) and in real headless Chromium with no host (`npm run ux:sweep`).
- 121 Excel-pane buttons, 26 PowerPoint-pane buttons, 49 Excel commands (17 ribbon + 42 shortcuts,
  shared handlers) and 14 PowerPoint ribbon commands pressed in four states each; 190 controls in the
  browser sweep (was 151: the 39 id-wired buttons were never pressed before).
- Found and fixed (v2.8.22): 16 Links-tab buttons dead without a supported Excel; Export / Push /
  Remove / Generate pressable mid-run (busy latch covered only `data-action` buttons); ribbon and
  shortcut errors toasting into a closed pane; copy buttons claiming success on a failed copy; the
  tool search pressing disabled and hidden buttons; the Inbox needing a manual refresh; Insert color
  key overwriting a block; three wrong sentences (doubled "Audit overlay" prefix, "1 matching cells",
  a blocked shortcut card reported as opened); Reveal left in the wrong state after any action.
- Simple to use (v2.8.23): two-click confirm on the eight destructive buttons, receipts that say
  what a format cycle did, honest labels (status words, "Pushed", "Update all" scope, Reset scope).
- Not seen in real Excel or PowerPoint: the desktop pass waits for Daniel's grant (section 5).

## 1. How every control was pressed

| Rig | What it proves | Result |
|---|---|---|
| `test/clickthrough.excel.integration.test.ts` (new, v2.8.21) | the shipped `taskpane.html` + `src/main.ts` over the strict fake Excel and a fake relay: every static and generated button pressed in four states (fresh workbook; a seeded model block with a chart and a second sheet; a multi-area selection; a protected sheet), every ribbon command with and without a selection; each press must react, answer a product sentence (never a bare office.js code, "Unknown action", "undefined"), raise no uncaught error and release the busy latch; coverage asserted against the HTML | green, 0 known defects carried |
| `test/clickthrough.ppt.integration.test.ts` (new, v2.8.21) | `pptpane.html` + `src/ppt/main.ts` over the fake PowerPoint and relay: all 26 buttons unpaired, paired, with two then three shapes selected (object tools succeed), and on a real inserted link row (update, revert, jump, change source, break); the search box and every filter; all 14 ribbon commands over a selection | green |
| `npm run ux:sweep` (extended, v2.8.21-22) | both panes in headless Chromium with no Office host: every button, tab, help toggle, select and checkbox clicked, incl. the id-wired ones; the toast is reset before each press so a repeated sentence still counts | 0 defects across 190 controls |
| `npm run ux:check` | layout at 320/360/420/500 px, four states per pane | 0 defects across 72 |
| `npm run check` | tsc, eslint, prettier, vitest, cargo, manifests, functions, shortcuts, version | green, 2905 tests (was 2863), zero skipped |

## 2. Excel pane, control by control

Verdicts: OK = pressed, reacts, answers as designed. FIXED = a defect found here and shipped.
CHANGED = a "simple to use" change shipped. DANIEL = needs his eyes or his call.

| Tab / section | Controls | Verdict | Note |
|---|---|---|---|
| Shell | 4 tabs, "Find a tool" search | OK / FIXED v2.8.22 | search no longer offers a disabled button (Copy report, Back) or the hidden Create/Cancel; glyph buttons found by their aria-label |
| Tools / New here? | Shortcut card, Got it | OK / FIXED v2.8.22 | a blocked pop-up now says "The shortcut card could not open. Allow pop-ups for this pane." |
| Tools / Selection inspector | Refresh, Shortcut card | OK | |
| Tools / Find a combination | Target, Tolerance, Find cells | FIXED v2.8.22 | "Found 1 matching cell" |
| Tools / Templates | 6 | OK | each writes its block, refuses a non-empty block |
| Tools / Model formatting | 5 presets, Clear, 4 number formats, Save 1-3, Use 1-3 | OK / CHANGED v2.8.23 | receipts name the step (see section 4) |
| Tools / Format cycles | 18 | OK / CHANGED v2.8.23 | each cycle now says the step it landed on; Row height / Column width say the size and "outside pls,fix Undo" |
| Tools / Model tools | Undo, Fill right/down, IFERROR, Autocolor, Insert color key, Consistent rounding, Unpivot, Comps stats, Mark, 4 pastes, 3 paste specials, x/÷1000, CAGR, ± sign, .00 +/- | OK / FIXED v2.8.22 / CHANGED v2.8.23 | Insert color key refuses a non-empty 6 x 2 block; IFERROR says added or stripped; CAGR says where it wrote |
| Tools / Audit | Audit overlay, Precedents, Dependents, Back, Select consistent region, Precedents of selection | FIXED v2.8.22 | protected sheet: "Audit overlay off: this sheet is protected, nothing was changed" (was doubled) |
| Tools / Charts | Waterfall, Tornado, Football field, Brand-format chart, CAGR label | OK | "CAGR label" acts on a number row, not a chart (label kept: DANIEL) |
| Workbook / Sheet explorer | Refresh, Unhide all, Show only this, Bury this, Move to end / up / down, Include very hidden, name + eye buttons per sheet | OK | "Move up/down" follow the vertical list; titles say left/right (kept) |
| Workbook / Super Find | query, 3 tick boxes, Find in workbook, result rows | OK | |
| Workbook / Styles | Scan styles, Delete N unused styles | OK | two-click confirm, now on the shared helper (v2.8.23) |
| Workbook / Table of contents | Insert contents sheet | OK | |
| Workbook / Share | Prepare for sharing, Clean past the data | CHANGED v2.8.23 | Clean past the data confirms on a second press (deletes rows and columns outside Undo) |
| Workbook / Model check | Run model check, Copy report, finding rows | FIXED v2.8.22 | a failed copy says "Copy failed: select the text and copy it by hand." |
| Workbook / Names | Scan broken names, Delete N broken names | OK | two-click confirm on the shared helper (v2.8.23) |
| Links / Export | Export selection, chart, table, text, chart picker | FIXED v2.8.22 | dead without a supported Excel: now "Excel is not connected." through the guard |
| Links / Linked objects | Project select, New project, Move to project, Name, Create, Cancel, Push selected, Push all, Go to source, Remove link, Auto-push, Highlight, row ticks | FIXED v2.8.22 / CHANGED v2.8.23 | busy latch now covers them; Remove link confirms on a second press |
| Links / Link key | Generate, Copy, Reveal, Forget | FIXED v2.8.22 / CHANGED v2.8.23 | Reveal follows whether a key exists after every action; Copy reports a failed copy; Generate (when a key exists) and Forget confirm on a second press |
| Brand / Palette | Reset, 7 pickers + 7 hex boxes | CHANGED v2.8.23 | Reset says "Brand settings reset" (it resets font, number style, currency and autocolor too) and confirms on a second press |
| Brand / Logo, Output defaults | Upload, Font, Language, Currency, Autocolor on edit, Copy palette JSON, Import JSON | OK | logo swatches need a decoded image: DANIEL in real Office |
| Brand / Keyboard shortcuts | Printable card, 42 key boxes, Apply, Reset all | OK / CHANGED v2.8.23 | Reset all confirms on a second press |
| Ribbon (17) + shortcuts (42) | all handlers | FIXED v2.8.22 | an error now shows the pane before its toast (success stays silent) |

## 3. PowerPoint pane, control by control

| Tab / section | Controls | Verdict | Note |
|---|---|---|---|
| Shell | 4 tabs | OK | |
| Links / Linked objects | Refresh, Update selected, Update this slide, Update all, Revert last update, Go to slide, Change source (+ list, Confirm, Cancel), Break link, search box, 4 filters, row ticks | OK / CHANGED v2.8.23 | filter options read like the badges ("Update available", "Source missing"); column "Pushed"; "Update all" title says it follows the filters; Break link confirms on a second press and says "The object stays on the slide." |
| Inbox | Got it, Refresh, Slide, Where, Paste latest linked, Insert per item | FIXED v2.8.22 | the Inbox reads the relay when its tab opens (no manual refresh) |
| Tools / Object tools + Smart Painter | Align (+mode), Distribute (+axis), Match size, Select similar, Swap, Capture, Apply | OK / CHANGED v2.8.23 | proven with 2 and 3 shapes selected; "centre" in the toast |
| Settings / Link key | key box, Save key, Forget key | CHANGED v2.8.23 | Forget key confirms on a second press |
| Ribbon (15) | all 14 commands + Links | FIXED v2.8.22 | an error shows the pane before its toast |

## 4. What shipped

- **v2.8.21** (harness): the two click-through suites, the sweep over id-wired buttons, an unknown
  `--pane` slug is an error.
- **v2.8.22** (fixes): items in the Outcome list; plus the sweep's toast sentinel.
- **v2.8.23** (simple to use): two-click confirm helper `src/ui/confirm.ts` on Clean past the data,
  Remove link, Generate (key exists), Forget key (both panes), Break link, Reset brand, Reset all
  shortcuts; receipts for the 16 cycles, IFERROR and CAGR; the label and sentence changes above.

## 5. Open, Daniel's side

- Real Excel and PowerPoint on the Mac: nothing above has been seen in a real host. The pass in the
  plan (v2.8.20 fix-wave behaviours, the PowerPoint crash matrix twice per route, the new confirms
  and receipts, shortcut card dialog, clipboard, logo upload, custom functions, Mac shortcuts) runs
  only on his explicit "go desktop" with both apps closed and him away from the machine.
- Windows pass, the M365 upload, the Cloudflare rate-limit rule: unchanged from `tasks/AUTORESUME.md`.
- Kept as they are, his call if he wants them changed: "Move up/down" for sheets (the explorer is a
  vertical list), "Pinstripes ↔" for columns, "CAGR label" under Charts, the 7 commands with no
  ribbon button (`PLSFIX_ROUND`, `PAINT_CAP1-3`, `TORNADO`, `UNPIVOT`, `SHARE`: a manifest change
  means the M365 re-upload), Ctrl+Shift+1/2/4/5/6/7 and Ctrl+Shift+F/A/O/V taking over Excel's own
  Windows shortcuts.
