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
- [ ] Tools > Find a combination: on the Variance sheet select B4:B17, type 2230 as Target,
      Find cells. Expect exactly four cells selected (B6, B10, B12, B16) and the line
      "4 cells selected · Sum 2 230 · Variance 0".
- [ ] Tools > Format: Save 1 on a formatted cell, Use 1 on another block; close and reopen the
      workbook, the three slot labels are still there (they travel with the file).
- [ ] Links > Export selection on P&L!B4:E9, Export as table on the same block, Export chart on
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
- [ ] Excel: change P&L!C4, Push all. PowerPoint: Update all. Every object refreshes in place,
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
