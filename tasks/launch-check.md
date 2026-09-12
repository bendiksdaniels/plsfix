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
