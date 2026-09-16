# Launch check on a real Office (Daniel, ~25 min)

v2.6.1 is live at dbautomatizacijas.com/modelis; the fakes and gates are green. This is the
one pass the fakes cannot replace: a real Excel and PowerPoint. Run it on the Mac first, then
once on Windows. Tick as you go; for any failure press the toast's **Copy details** and paste
the text back to me, that is the whole bug report.

Setup: quit and reopen Excel and PowerPoint (the wef folders carry `manifest.prod.xml`; the
ribbon reloads only on a fresh launch). Open `demo/out/pls,fix Demo Model.xlsx` (`npm run
demo:build` if missing) and a blank deck.

## Excel

- [ ] Ribbon: the pls,fix tab shows five groups, every button with its own icon.
- [ ] Tools > Find a combination: on the Variance sheet select B11:B24, type 2230 as Target,
      Find cells. Expect exactly four cells selected (B13, B17, B19, B23) and the line
      "4 cells selected · Sum 2 230 · Variance 0".
- [ ] Tools > Format: Save 1 on a formatted cell, Use 1 on another block; close and reopen the
      workbook, the three slot labels are still there (they travel with the file).
- [ ] Links > Export selection on P&L!B11:E16, Export as table on the same block, Export chart on
      the Revenue chart, Export as text on one cell. Each shows a green toast within seconds.
      A "GeneralException" toast here is the Mac bug: Copy details and paste it.
- [ ] Links > Export the line chart on the Bridge sheet (Revenue trend) as a chart.

## PowerPoint

- [ ] Ribbon: three groups, Links, Objects (six buttons), Arrange (eight buttons), all with
      icons.
- [ ] pls,fix > Links > Settings: paste the link key from Excel; Inbox lists the exports.
- [ ] Inbox > Paste latest linked: the newest export lands on the active slide in one click.
- [ ] Insert the table, the text and both charts. The column chart and the line chart arrive
      as groups of native shapes with value labels (the line one: markers, connectors, a rising
      segment drawn correctly, labels above the markers); the table is a native table.
- [ ] Links list: type "Bridge" in the search box, then try the Source workbook, Slide and
      Link status filters; a ticked row stays ticked when a filter hides it.
- [ ] Excel: change P&L!C11, Push all. PowerPoint: Update all. Every object refreshes in place,
      sizes kept.
- [ ] Draw three rectangles. Tools tab: Align left, Distribute across, Match size, Select
      similar, Swap (two selected). Then the same from the ribbon's Arrange and Objects groups.
- [ ] Smart Painter: give one rectangle a fill and no outline, Capture, select the others,
      Apply. Fill copied, no outline written, no error.
- [ ] Ribbon > Objects > Object tools: the pane opens on its Tools tab.

## Windows (once)

Same list; the point is the ribbon icons, the shared runtime commands and the chart groups on
a second platform. Note the Office build (File > Account) with the result.

## Added by the 09.09 audit wave (v2.6.3 to v2.6.15): only a real Office can show these

The fakes proved every fix; these are the behaviours the agents could not make a fake
refuse or answer. Same rule: a wrong one gets the toast's **Copy details** pasted back.

Excel

- [ ] Ctrl+Space a whole column, Autocolor: "Autocolor supports up to 5,000 selected cells
      at once." at once, no wait. Ctrl+A the whole sheet, then a number format or any cycle:
      the cap sentence, no freeze (Range.cellCount answers -1 there).
- [ ] Whole column, Title preset: every row becomes 25 pt and pls,fix Undo cannot put the
      heights back. Decide: acceptable, or should the preset skip the height above the cap?
- [ ] Row height on a selection that includes a hidden row: the ladder restarts at 15 pt and
      the row shows. Row height on a protected sheet: "row height: this sheet is protected,
      nothing was changed", not Excel's own string.
- [ ] Mac: Date cycle on a date cell, pressed repeatedly: dd.mm.yyyy, mmm-yy, yyyy, back.
- [ ] Restyle a pie on desktop 365: labels outside the slices, no leader lines any more (the
      label-level property is ExcelApi 1.19; the series-level one, 1.9, would bring them back).
- [ ] A template or the tornado with the active cell in the sheet's last rows: a raw
      InvalidArgument is expected (the fake has no grid edge). Restyle a 3-D clustered column:
      OutsideEnd labels may be refused. Unpivot with the workbook structure protected: Excel's
      own AccessDenied string. NPV/IRR template with a positive outlay: INDEX(cumulative, 0)
      may spill under the block. A chart restyle is one round trip longer on the web.
- [ ] Latvian or Russian desktop Excel, Brand tab: the note under Output defaults reads
      "Excel shows 1 094 417.5, the house style." and names no setting to change.
- [ ] Prepare for sharing on a workbook with a protected sheet whose "Select locked cells" is
      off: expect a clean run or one raw Office error (range.select on every visible sheet).
- [ ] Excel for the web on a read-only or co-authored workbook: drag a palette colour and
      watch for a red toast per step.
- [ ] Push a link whose source sits on a hidden sheet: a picture, or one per-link failure
      line. Go to source on a chart dragged to another sheet: the sheet activates and the
      chart is selected. The chart picker after a chart export lists the PLSFIX_LINK_ anchor
      name as an option: internal bookkeeping in a dropdown, a UX call.
- [ ] Export a range that seals past 4 MiB and push it: "That export is too big to send.
      Export a smaller range."

PowerPoint

- [ ] Pull the network in the middle of Update all: within about a minute the toast says
      "PowerPoint stopped answering while repainting the links" and the pane is usable again
      (before: busy for ever, pane reload the only heal). Same on an insert, an object tool
      and a scan of a deck with groups.
- [ ] Insert onto a slide whose layout has a picture or table placeholder with no text: the
      insert lands normally (placement reads textFrame.hasText on every placeholder).
- [ ] Update a text link whose box you restyled by hand: the first run's formatting may
      spread over a part-coloured number.
- [ ] Export a pie with a negative slice: Excel plots it at its absolute size, the shapes
      skip it; compare the picture and the group on one slide.
- [ ] Busy web slide: insert the demo pie (21 shapes) and the column chart (20). Neither
      may show "as a picture: PowerPoint stopped answering while drawing the shapes"; if a
      healthy chunk trips the 60 s deadline, the constant is too tight. Force a jam and check
      afterwards that the swallowed chunk's shapes never appear late.
- [ ] A chart with an empty category name draws an empty text box.
- [ ] Excel, cold launch: press a pls,fix shortcut or ribbon button before ever opening the
      pane. It must work (the commands register at page load now, as PowerPoint's do).
- [ ] Smart Painter: Capture from a solid shape, Apply onto a native table's outer border
      and onto an empty content placeholder: paints, or refuses in a sentence, never a raw
      error. Ribbon > Objects > Object tools when the pane cannot open: nothing is said.
- [ ] Update a formatted table link (the demo P&L, 6x4) after a source edit: "1 updated" and
      the row "Up to date" within about 15 s on desktop, under a minute on the web, never
      "stopped answering while repainting the table" (cells go eight per round trip now; on
      the web a formatted cell cost about 0.4 s per property on 09.09).

## Added by the v2.7 wave (12.09): the comps and valuation tools

- [ ] Comps stats: on a real comps table, the six formulas recalculate when a multiple is edited,
      and `=PERCENTILE.INC` resolves (a non-English Excel may expect the localised function name).
- [ ] Comps stats: the written block's number formats match the source column in a currency and in
      a multiple format, on Excel for Mac and Excel for the web.
- [ ] Football field: the stacked bar really shows floating bars in desktop Excel - the floor series
      invisible, no outline artefact at its top edge - and the legend is off.
- [ ] Football field: `categoryAxis.reversePlotOrder` puts the first selected method at the top in
      desktop Excel and on the web (the fake only records the flag).
- [ ] Football field: `valueAxis.numberFormat` is honoured on Excel for the web (axis number format
      is ExcelApi 1.8; the web build has been known to ignore chart axis writes).
- [ ] Football field: the chart lands beside the helper block, not on the model, on a real sheet
      with non-uniform row heights.
- [ ] Pinstripes: `#E8E9EC` is visible but not distracting on a real screen and when printed, and
      it reads correctly against a workbook whose own theme is dark.
- [ ] Pinstripes: a second press over a real hand-formatted grid clears exactly the bands and
      nothing else - Excel's own read-back of a solid fill must return the same hex it was given
      (see the 2026-08-27 lesson on read-back canonicalisation).
- [ ] Pinstripes over columns on a sheet with grouped or hidden columns: the bands follow the
      selection's own column order, not the visible one.
- [ ] All three: the toasts fit the pane at a 320 px dock without wrapping past two lines.

## Added by the v2.7 wave (12.09): the hygiene tools


- [ ] Excel desktop: the Indent, Align and Underline buttons step a real selection and, pressed
      again, keep stepping - the read-back Excel gives (`SingleAccountant`, `General`, a null
      indent on a mixed range) lands on the rung the fake predicts.
- [ ] Excel desktop: pls,fix Undo puts back an indent, an alignment and an underline the cycles
      overwrote (the fake proves the capture; only Excel proves `setCellProperties` restores all
      three).
- [ ] Excel desktop: Unhide all, Show only this, Bury this and the three moves against a real
      multi-sheet workbook, including a VBA-set `xlSheetVeryHidden` sheet, and Bury on the last
      visible sheet answering "Excel needs one visible sheet."
- [ ] Excel desktop: with Review > Protect Workbook (structure) on, each of the four sheet tools
      answers "this workbook's structure is protected, nothing was changed" and the tab strip is
      unchanged - confirms Excel really returns AccessDenied there and `workbook.protection`
      reads true.
- [ ] Excel desktop: Clean past the data on a real workbook whose used range runs to row 60 000
      shrinks it (Ctrl+End lands on the data afterwards, and the saved file gets smaller).
- [ ] Excel desktop: Clean past the data on a sheet with a chart keeps every row and moves no
      chart, and the toast names the reason.
- [ ] Excel for the web: Clean past the data on a large sheet finishes inside the pane's patience
      (one delete per band, not per row).
- [ ] Excel desktop and web: `=PLSFIX.CAGR(100,200,4)` returns 0.189207 in a cell, the function
      appears in the formula autocomplete with its three argument names, and its help link opens
      the support page (needs slice K3's page to exist).

## Added by the v2.7 wave (13.09): PowerPoint and shell follow-ups


- [ ] Excel -> PowerPoint: export a very tall, narrow chart (a sheet chart about 200 x 2000 px, e.g. a
      40-row bar chart squeezed into two columns). It must arrive as the picture with "as a picture: the
      chart would be smaller than 200 x 120 pt on this slide", never as a group of hairlines.
- [ ] PowerPoint: insert a bar chart of 30 to 40 rows whose values are nine digits. Every value label
      must sit inside the chart's own box - the longest bar's label ending at the bar's tip, right
      aligned - and the group's box must be the one the placement gave it.
- [ ] Both panes, keyboard only: Tab into the tab strip (one press), walk it with the arrows, Home and
      End, then Tab again - the next press must leave the strip, not walk the remaining tabs. Landing on
      the Excel pane's Workbook tab by arrow must show freshly read sheets, exactly like a click.
- [ ] Excel, pane closed, cold launch: Ctrl+Shift+Alt+F (Find in workbook) and Ctrl+Shift+Alt+S (Scan
      unused styles) must each open the pane on the Workbook tab, with the caret in the box / the style
      list showing, not run silently.
- [ ] PowerPoint: a chart link pushed from an Excel build whose chart data this pane cannot read (a
      newer schema, or a hand-edited relay blob) must still insert - the picture with "as a picture:
      chart data unreadable" - and Update all must keep repainting it. Only a real cross-version push
      exercises this; the fakes cannot.
- [ ] PowerPoint: insert a chart onto a slide already holding a large shape: the group must arrive full size (never shrunken below 200 x 120 pt) and the pane says "Placed over other objects".
- [ ] PowerPoint: a stacked bar chart with 3 and with 6 series: one row per category, segments side by side, nothing outside the group box.

## Added by the v2.7 wave (13.09): relay timeouts and the store-valid manifest


- [ ] A genuinely dead or very slow relay (block the host at the network level, or point
      MODELIS at an unreachable address) shows "The link relay did not answer in time."
      within about 20s in both the Excel and PowerPoint panes, and the pane recovers
      (no permanent "busy") rather than hanging.
- [ ] Sideload `manifest.prod.xml` and open Office's "My Add-ins" management dialog (or
      the AppSource-style install card) to confirm the 64px icon actually renders at
      high DPI, since only Office's own UI (not `office-addin-manifest validate`) draws
      it.
- [ ] After deploy, `https://dbautomatizacijas.com/modelis/support.html` and
      `.../privacy.html` load through the live Cloudflare Access bypass path exactly as
      `taskpane.html` does today (the dev worktree only proves the route logic, not the
      live edge).
- [ ] A real, slow-but-not-hung SQLite write under production load (not this slice's
      deliberately-`tokio::time::sleep`-ing test handler) still gets a 408 rather than
      holding the worker - `rusqlite` is a blocking driver, and a genuinely synchronous
      hold inside a handler is not preemptible by `TimeoutLayer`'s `tokio::time::sleep`
      race the way an `.await`-yielding handler is. Worth a real load test before
      leaning on this for abuse protection rather than just slow-client protection.

## Added by the v2.7 wave (13.09): the audit tools


- [ ] Excel desktop: with a filled model row selected, "Select consistent region" selects exactly the
      filled block and the toast counts it; the same on a block filled both ways.
- [ ] Excel desktop: "Select consistent region" inside a real model sheet (used range far past 5 000
      cells) grows normally, and only refuses when the block around the active cell is itself over
      5 000 - and then says "Select consistent region supports up to 5,000 cells at once."
- [ ] Excel for the web: "Precedents of selection" over ~20 formula cells answers in one visible
      round trip (no per-cell stutter), and the union really becomes the selection.
- [ ] Excel for the web / desktop: a selection where two or three cells are constants
      (`=1+2`, a typed number) still answers, with the halve-then-flat retry costing seconds not
      minutes, and the toast naming the cells with no formula.
- [ ] Excel desktop: a precedent on another sheet is listed as a chip and clicking it jumps there,
      while the initial select stayed on the reviewed sheet.
- [ ] Excel 2019 (below ExcelApi 1.12): "Precedents of selection" says "Tracing needs a newer Excel
      build." and nothing else happens.
- [ ] Pane at a 320px dock: the two new rows and the grouped list read cleanly, group labels on their
      own line (`.chip-group` really does break the flex row in the Office webview).
- [ ] Excel desktop: a sheet protected with selection disabled
      (`worksheet.protection.protect({ selectionMode: Excel.ProtectionSelectionMode.none })`) - both
      read-only tools end on `.select()`, and the host may refuse that with a raw string instead of a
      pls,fix sentence. Shared with `src/excel/reconcile.ts`, which selects the same way; if it does
      refuse, all three need the same staged wording.

## Added by the v2.7 wave (13.09): the shortcut manager

- [ ] Excel desktop (Windows, Microsoft 365, signed in): Brand > Keyboard shortcuts fills the boxes that already carry a custom key on that account, Apply toasts "Shortcuts updated for your account.", and the remapped key runs the action in the sheet.
- [ ] Excel desktop (Mac, signed in): the rows print `Cmd`/`Option`, and the key that actually fires the action is confirmed - Ctrl or Cmd - since the documentation only states Cmd -> Ctrl on Windows and Alt -> Option on Mac, never Ctrl -> Cmd on Mac.
- [ ] Excel for the web (signed in): the panel works and the remapped shortcut fires with focus on the grid (the docs warn custom shortcuts do not fire while the task pane has focus).
- [ ] Signed OUT of Office (or a local/anonymous account): Apply toasts exactly "Custom shortcuts need Microsoft 365 with a signed-in account.", and the copyable details block carries the real `code:` Office rejected with - record that code so the guess in section 3 can be replaced by the documented one.
- [ ] Excel too old for KeyboardShortcuts 1.1 (Windows before 2111 / Mac before 16.55): the section shows the same sentence with every box disabled, and neither button throws anything else.
- [ ] Deliberate clash: set a pls,fix action to a combination another add-in or Excel already owns; Apply toasts "... Already used elsewhere: <key>." and Excel's own conflict dialog appears on the first press of that key.
- [ ] Reset all on an account with several custom keys: every key goes back to the shipped default, the boxes empty, and reopening the Brand tab still shows them empty.
- [ ] The keys survive: sign in to Excel on a second machine with the same account and confirm the custom keys roamed (the docs say roaming settings, per platform).
- [ ] Printable card from the section heading opens the same dialog as the Tools tab's card button.
---
The Mac row in section 8 is replaced by these two:
- [ ] Excel desktop (Mac, signed in): a `Ctrl+...` row fires on the physical key the Mac reports, and the pane prints `Ctrl`, not `Cmd`. Record which physical key it actually is: the documentation states only Cmd -> Ctrl on Windows and Alt -> Option on Mac, so this is the open question.
- [ ] Excel desktop (Mac, signed in): set one action to a `Cmd+Shift+<key>` combination and confirm Office accepts the string and the Command key fires it, since the docs say Cmd is supported on macOS but no sample sends it through `replaceShortcuts`.

## Added by the v2.7 wave (13.09): the paste suite

- [ ] Duplicate Formulas paste on a real model: copy a 3-column block whose formulas mix
      in-block references and absolute references to an assumptions cell, paste it four rows
      down, and confirm Excel recalculates to the same numbers the source block shows.
- [ ] Duplicate Formulas paste onto a DIFFERENT sheet, now that outside references are
      qualified: confirm Excel accepts `=Model!Z9` and `='P&L 2025'!Z9` as written and the
      pasted block returns the same numbers as the source block.
- [ ] Duplicate Formulas paste from a sheet whose name Excel would quote (`P&L 2025`, `Q1`,
      `Bob's`) onto another sheet: confirm no `#NAME?` and no repair prompt on reopen.
- [ ] Duplicate Formulas paste over a source containing an array/dynamic formula and a
      `LAMBDA`/`LET` name: confirm Excel accepts the rewritten text without a spill or name error.
- [ ] **Real Excel's `copyFrom` tiling vs `tileGrid`'s modulo:** paste formats (the existing
      button) and then number formats only (the new one) into a destination that is NOT a whole
      multiple of the source - a 2x2 source into a 3x3 selection - and confirm both leave the
      same formats in the same cells. `tileGrid` repeats by modulo, cutting the last tile short,
      which is what the fake host's `copyFrom` does; real Excel may instead paste the source
      once at the top-left or refuse the shape, and if it does, `planFormatWrites` is the place
      to match it.
- [ ] Paste number formats only onto cells carrying conditional formatting and a table style:
      confirm the number format lands and neither the fill nor the table banding changes.
- [ ] Paste number formats only over a merged block in real Excel: confirm the format applies and
      Excel does not refuse the write the way it refuses a value write across part of a merge.
- [ ] Paste row heights only with an autofit row and a wrapped-text row in the source: confirm
      the target rows take the source's rendered heights and do not re-autofit on the next edit.
- [ ] Paste row heights only on a sheet with a frozen pane and a filter-hidden row: confirm a
      filter-hidden TARGET row is written and a filter-hidden SOURCE row is treated as hidden and
      skipped, which is what `rowHidden` reports.
- [ ] Paste row heights only on a protected sheet whose protection granted `allowFormatRows`:
      confirm Excel accepts the write instead of returning AccessDenied (the fake refuses it
      either way).

## Added by the v2.7 wave (13.09): the Excel follow-ups

- [ ] Fast fill's new refusal ("Fast fill supports up to 5,000 cells at once.") shows as a clean pane toast on a real 6,000+ row/column fill, not a raw host timeout, on both desktop Excel and Excel for the web.
- [ ] Super Find and Prepare for sharing on a real ~200,000-cell sheet (a genuine deal model, not the fake's bounding-box shortcut) complete within the pane's usual response time and do not trip the host's own `RequestPayloadSizeLimitExceeded` the old, lower total cap existed to avoid.
- [ ] `Worksheet.names` (ExcelApi 1.4) behaves on a real host exactly as the fake now models it - add a name scoped to one sheet in real Excel, break it (delete what it points at), and confirm the Workbook tab's scrubber lists it as `Sheet!Name` and deletes it.
- [ ] A real desktop-365 pie chart (ExcelApi 1.9 available, 1.19 not yet) visibly shows leader lines after "Chart Smart Format" - this is the entire point of the series-level fix and only a real host can confirm the line actually renders, not just that the property was set.
- [ ] The "Linked chart · xxxx" label reads well at the narrowest supported pane width (320px) with a real, longer chart list in the Links tab dropdown (ux:check only exercises synthetic fixture data).
- [ ] Excel desktop: delete broken names on a workbook whose structure is protected: the toast says "The workbook's structure is protected, so nothing was deleted." and no name is gone afterwards (Office.js may have applied the first deletes of the batch before the refusal; if any name is missing, the sentence must change).

## Added by the v2.7 wave (13.09): the comps polish

Host-only behaviours the fakes cannot confirm (P1, v2.7.10):

- [ ] Comps stats over a comps table with hidden rows inside the selection: the six formulas span the whole block (hidden rows included) and the block lands under the last row, hidden or not.
- [ ] Comps stats under a filtered table: the block lands under the table's last row, not under the last visible one, and the statistics count every filtered-out row.
- [ ] Comps stats over a column a modeller formatted by hand (a text or date column beside the multiples): its format survives the block.
- [ ] Football field with hidden columns between the label and the numbers: the helper block still lands in the first three columns right of the whole selection.
- [ ] Football field on Excel for the web, and on a build below ExcelApi 1.7: row 1 on top with `reversePlotOrder`, and the "plain axes on this build" note when the host has neither 1.7 nor 1.8.
- [ ] Football field on a real protected sheet: the pane says "this sheet is protected, nothing was changed" and the Undo row still names the previous action.
- [ ] Pinstripes over a range with hidden rows: Excel bands the hidden ones too, and unhiding shows the band (decide whether that is what we want).
- [ ] Pinstripes over an Excel table with its own banded-row style: our tint over theirs, and the second press still clears ours.
- [ ] Pinstripes on a protected sheet with an unlocked island: the note, and nothing painted.
- [ ] Comps stats, the football field and the tornado on a protected sheet whose target block is an unlocked island: all three now refuse with "this sheet is protected, nothing was changed" where before they would have written (house-consistent with Autocolor and the audit overlay; a user of protected templates will notice).
- [ ] The Charts section reading order in a real docked pane at 320 px: waterfall, tornado, football field, brand-format, CAGR, no wrapping or clipping of the moved button's label.

## Added by the v2.7 wave (13.09): the hygiene stress pass

Host-only behaviours the fakes cannot settle (P2, v2.7.11):

- [ ] Ctrl+A on a sheet, then Indent / Align / Underline / Fill / Font colour: Excel applies one format property over the whole sheet without freezing the pane, and the toast reads "Selection updated (too large for undo)".
- [ ] Ctrl+A on a sheet, then Row height: `range.getEntireRow().format.rowHeight` over 1 048 576 rows returns, and the toast is a sentence rather than a host string.
- [ ] A row-height cycle whose band starts on a HIDDEN row: Excel's `format.rowHeight` for a hidden row (0 or its stored height) decides whether the press unhides the band; the fake cannot model it.
- [ ] A protected sheet with SOME cells unlocked, ctrl-click one unlocked and one locked cell, press Indent: check whether real Office.js applies the unlocked area before rejecting the batch (the skipped P2 row assumes it does).
- [ ] Clean past the data on a sheet whose merge straddles the data edge: Excel shrinks the merge rather than refusing the row delete, and the top-left value survives.
- [ ] Clean past the data twice on a sheet with a chart: real Excel does not shrink the used range when only formats are cleared, so the second press should say "Cleared the formats past the data" again (the fake says "Nothing past the data").
- [ ] `=PLSFIX.CAGR(A1, B1, C1)` where A1 holds an error cell (`#N/A`): confirm Excel propagates the error to the cell rather than calling the function.
- [ ] `=PLSFIX.ROUNDSUM(A:A, 0)` on a whole-column reference: confirm Office marshals it and the cap sentence is what the cell shows.
- [ ] Bury this sheet on the active sheet: confirm Excel activates another sheet (the explorer keeps the buried sheet on its list, muted; that half is in the suite).

## Added by the v2.7 wave (13.09): the relay hardening

Against the deployed relay through the real tunnel and unit (S1, v2.7.12):

- [ ] With `MODELIS_TRUSTED_PROXY=cloudflare` live, two people on different networks each push a link in the same minute and neither sees a 429 (the bucket is per client through the tunnel, i.e. nginx passes `CF-Connecting-IP` through to :8804).
- [ ] A normal working session (a dozen exports and a chart picture or two) never sees a 429 from the byte budget.
- [ ] A real "Push all" / "Update all" over a deck with more than 50 links does not report 429s at 100 writes per minute (the pane does not retry one); if it does, raise `MODELIS_MAX_INFLIGHT_WRITES` to 64 and `MODELIS_RATE_WRITE_PER_MIN` back to 300 in the unit, never lengthen the body deadline.
- [ ] `curl https://dbautomatizacijas.com/modelis/version` still parses in the gateway's status aggregator and the dashboard, and its `relay` numbers move within 30 s of a push.
- [ ] A deck's "Update all" over a slow link still gets its 304s (the `no-store` change did not break the ETag round trip in a real Office webview).
- [ ] `journalctl -u` on the pane host's unit shows the startup line naming the trust mode, no "every client shares one bucket" warning, and no `relay rate write ...` lines during normal use.

## Added by the v2.7 wave (13.09): the PowerPoint stress pass

Host-only behaviours the fakes cannot settle (P3, v2.7.13):

- [ ] PowerPoint (desktop and web): delete a linked picture, then press Update all before the list refreshes: the toast names the link and says the object is no longer where the list had it, never `ItemNotFound`. The same for Break and for Revert on a row whose shape was deleted.
- [ ] Go to slide on a row whose slide was deleted: does the real host throw? The fake accepts any slide id. If it throws `ItemNotFound`, wire `missingShapeError` into `goToSlide`.
- [ ] Smart Painter: capture from a native PowerPoint table and from a group on a real host (`Shape.fill.type` may be unreadable there, where the fake answers Solid/NoFill).
- [ ] PowerPoint for the web: a table rebuild (source grew a row) keeps the row on "Update available" until the last format chunk lands, and the next press finishes the formats.
- [ ] A 60 x 20 table repaint on the web: confirm the 150 round trips complete.
- [ ] An object tool against a shape a protected layout or a locked group refuses to move: the host's own refusal string reaches the toast; check it reads as a sentence on a real host.
- [ ] Press Enter in the link-key field while Update all is running: "Wait for the last action to finish." and no button comes back before the batch does.
- [ ] Press a ribbon object tool (Align left) while Update all is running: the same sentence, and the tool works on the next press.
- [ ] Insert a 61 x 21 table export (a hand-written relay row or an older Excel build): "Tables go up to 60 rows and 20 columns; export a picture for more." and nothing on the slide.
- [ ] A relay that answers 429 or 507 on the real deployment: the toast reads "The relay is busy..." / "The relay is full...", never the status line.
- [ ] An insert whose `deleteInbox` the relay refuses (500): "Inserted ...", one shape, and the Inbox no longer offers that export in this pane session.
- [ ] Select similar refuses with "not one inside a group": confirm on a real host that nothing else (a shape on a layout, a placeholder reached another way) can be selected while sitting outside `Slide.shapes`; the guard detects the class, the sentence names the one cause.

## Added by the v2.8 wave (13.09): guided demos and placement

- [ ] The guide band ("Try on this sheet" plus a Commands line, rows 1-7) shows on every
      visible sheet of `pls,fix Demo Model.xlsx`; the P&L block the Excel and PowerPoint
      sections above export and edit shifted down 7 rows with it - read `P&L!B4:E9` there as
      `P&L!B11:E16` and `P&L!C4` as `P&L!C11`; read the Variance column B4:B17 in the rows
      above as B11:B24, and the four cells B6, B10, B12, B16 as B13, B17, B19, B23.
- [ ] Open `demo/deck/pls,fix Demo Deck.pptx` (rebuild with `npm run demo:deck`) and its four PowerPoint
      tasks: each placement lands exactly where chosen (the P&L table on slide 2, Whole slide;
      the Revenue chart and the Segment pie on slide 3, Left half and Right half; a Data picture
      on slide 4, Selected shape). The empty placeholder on slide 4 is gone once the picture is
      in; the two dashed rectangles on slide 3 line up with the Left half / Right half spots.
- [ ] Paste latest linked honours both pickers too: set Slide to a chosen slide and Where to a
      spot other than Free space, then Paste latest linked - the newest export lands there, not
      on the active slide in free space.
- [ ] With Slide left on "This slide" and Where left on "Free space" (the defaults), every
      Insert button and Paste latest linked behave exactly as they did before v2.8.0.
- [ ] With Where set to Selected shape and nothing selected on the slide, the pane says "Select
      a shape on the slide first, or choose another spot." and Insert does nothing.
- [ ] Update all after inserting through the new pickers: objects refresh in place exactly as
      before, position and size kept.
- [ ] Insert (or Paste latest linked) onto a slide other than the one showing: the pane
      switches to that slide afterwards.
- [ ] Open a deck with more than a few slides: the Slide picker lists every one of them
      (plus "This slide" first), not just the ones visited this session.

## Desktop pass on the Mac (13.09, v2.8.1; Excel and PowerPoint 16.107, macOS 26.6)

Run with computer use on Daniel's own Excel and PowerPoint, the demo workbook and the deck copy.

- [x] Excel: the pls,fix tab and Model Tools open the pane ("Excel connected"); Autocolor on
      `P&L!A10:H27`; the audit overlay marks F16 and G19 and clears again; Links exports the P&L
      table (B11:E16), the Revenue chart, the Segment pie and a Data range: four rows under
      Linked objects, each "Pushed just now".
- [x] PowerPoint: the sideloaded add-in loads from Home > Add-ins > pls,fix (once per launch);
      pairing with the copied key; the P&L table lands on slide 2 with Slide = Slide 2 and
      Where = Whole slide (natural size, centred: fitInto never scales up); a Data range picture
      lands in free space on the active slide.
- [x] FAILED on v2.8.1: the Revenue chart (Slide 3, Left half) and the Segment pie (This slide,
      Free space) each crashed PowerPoint for Mac right after the draw batches returned; the deck
      came back as [Autosaved] without the chart and, once, unpaired. v2.8.2 kept every chart a
      picture on the Mac for one release; the bisection on the Mac (13.09 evening) found the
      shape: one `addGroup` of a whole chart (19 shapes) dies, six shapes group fine, the same
      shapes grouped through the ribbon are fine. v2.8.3 groups in sub-groups of six on the Mac
      (`groupTier`), proven there: insert, a redraw through Update all, both alive.
- [x] v2.8.2 on the Mac: the two halves on slide 3 (as pictures then), the placeholder on slide 4
      (consumed, the picture at its box), Paste latest linked with Slide 2 + Top right, Update all:
      all green.
- [x] v2.8.3 on the Mac with the released pane (13.09 19:20): the Revenue chart (Slide 3, Left half)
      and the Segment pie (Slide 3, Top right) land as shape groups, PowerPoint alive; ungroup once
      shows the sub-groups.
- [ ] Open, cosmetic: at a quarter-size spot the pie's value labels and legend entries wrap
      ("1,51 / 9", "Retai / l"): the label boxes are laid out for the plan's size and then scaled
      with the box, the font is not. Halves and the whole slide are fine.
- [ ] Open, placement: Free space on a slide with no free spot falls back to a centred full-size
      overlap (by design, with the note), and on the Mac a second chart landed exactly over the first
      (13.09 19:00) and a third inside a dashed frame (19:06): the scan did not count the groups as
      occupied. Next session, item 3 of the NEXT SESSION brief in `tasks/AUTORESUME.md`.

## Added 16.09: the table header

Against a real Excel and PowerPoint: the fakes prove the logic, not the deck.

- [ ] Export a table whose first row is entirely bold: the inserted table shows PowerPoint's
      own header row (a shaded band across row 1), not a plain grid.
- [ ] Update all on that table: the header band and the deck's own table style are still
      there afterwards, exactly as before the update.
- [ ] Clear a cell's fill in Excel and Update all: the matching cell in the deck loses its
      colour too, and every other cell's format is untouched.
- [ ] Export a range whose first row is not entirely bold: the inserted table shows no
      header band.


## Desktop pass on the Mac (16.09, v2.8.9 to v2.8.13; Excel and PowerPoint 16.107, macOS 26.6)

Driven from the terminal (screencapture + CGEvent clicks, no computer-use grant needed) on the demo
workbook and a scratch copy of the demo deck; every row below was seen on screen.

- [x] Excel v2.8.12: Autocolor leaves labels alone and colours the growth row black (identity constants);
      the audit overlay tints the typed G19 solid and no longer stripes H19; Fill formula right keeps the
      source's number format; ribbon Precedents on D17 selects D14 and D16 together; ribbon CAGR writes
      10.0 % beside the CAGR sheet's first row; the Header preset bands the P&L header row.
- [x] Links: Export as table / active chart / selection / as text all push; New project + Move to project
      work; the PowerPoint Inbox groups the exports under the project; Push all then Update all repaints
      every link (7 updated); Revert last update repaints the older revision; Break link works.
- [ ] PowerPoint v2.8.18: the Linked objects list opens as folders ("AMASTY · 2", then "NO PROJECT")
      with the workbook and kind under every object at the pane's default width; ticking, Update
      selected and Go to slide still act on the rows, never on a folder header.
- [x] PowerPoint v2.8.9: the Revenue column chart (left half), the Segment pie (right half), the EBITDA
      margin line chart (free space) and the 40-point Big40 chart (slide 1, free space) all land as shape
      groups; a text link and a placeholder picture land on slide 4; PowerPoint alive after every insert and
      after Update all (process start time unchanged).
- [x] FAILED, fixed in v2.8.13: a freshly inserted table came double height (rows sized for 18 pt); numbers
      read "12,400" where Excel shows "12 400"; an export made with the overlay on shipped its tint.
- [x] Table header band on the Mac (v2.8.16): the pasted table shows the band, `firstRow="1"` + the
      style id in the saved XML, rows at the payload's 11 pt, "13 500" separators.
- [x] v2.8.14 on the Mac: the pie's labels are single-line; a Free-space insert on a full slide answers
      "Placed over other objects: the largest free spot is 888 x 69 pt" instead of a sliver. Still open:
      the slide-1 text-extent case (a chart below a short list) and a re-inserted 40-point chart.
- [x] pls,fix Undo after x1000 (v2.8.15): the sanitised restore; proven in the fake with the Mac's own
      snapshot shape, not yet re-pressed on the Mac after the deploy.
- [x] Custom functions (v2.8.17): a typed =PLSFIX.CAGR(100,161.051,5) gives 0.1 on the live pane; the
      rebuilt demo's CAGR sheet computes all six rows at load (two refused, as designed).
