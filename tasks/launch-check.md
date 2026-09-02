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
