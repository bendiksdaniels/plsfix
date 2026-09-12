# Notes for certification (paste into Partner Center)

pls,fix is a task-pane add-in for Excel and PowerPoint with add-in commands (a "pls,fix"
ribbon tab in both hosts), Excel custom functions, and a link relay between the two hosts.
No account, no sign-in, no purchase. Everything below works on Excel and PowerPoint for
Windows, Mac and the web with a Microsoft 365 subscription.

## Setup (2 minutes)

1. Sideload `manifest.prod.xml` (the package of this submission) into Excel and into
   PowerPoint, or install from the listing once published.
2. Download the test workbook: https://github.com/bendiksdaniels/plsfix/releases/latest/download/pls.fix.Demo.Model.xlsx
   and open it in Excel. Its **Start here** sheet is a checklist of every tool; the steps
   below are the short path.

## Excel: pane and commands (5 minutes)

1. Ribbon tab **pls,fix** > **Open pls,fix** (or Ctrl+Shift+M): the pane opens on the Tools
   tab; the status pill reads "Excel connected".
2. Sheet **P&L**: select B4:E9, click **Autocolor selection**: hardcoded numbers turn blue,
   formulas black. Click **Toggle audit overlay** (Ctrl+Shift+A): consistent formula rows are
   striped and the two planted errors turn soft red; click again to restore the fills.
3. With any cell selected, press Ctrl+Shift+1 three times: the number format cycles
   (general, two decimals, thousands).
4. Tab **Workbook** > **Model check**: a list of findings with a jump to each cell
   (two formula inconsistencies, one broken name, one hidden sheet in the test workbook).
5. Tab **Tools** > Templates > **DCF**: select an empty cell first; a formatted calculation
   block is written at the selection.

## Excel: custom functions (required test)

1. Sheet **Rounding** holds five numbers in B4:B8 whose exact total is in B9.
2. In C4 enter `=PLSFIX.ROUND($B$4:$B$8, 1, 0)`, then fill down to C8 (the second argument is
   the position in the range, 1 to 5; the third is the decimals).
3. In C9 enter `=PLSFIX.ROUNDSUM(B4:B8, 0)`.
4. Expected: C4:C8 are whole numbers that add up exactly to C9, and C9 equals B9 rounded to
   a whole number, even where plain ROUND of each cell would not add up. The **Consistent
   rounding** button in the Model tools list writes the same formulas beside the selection.

## Excel -> PowerPoint links (5 minutes)

1. Excel, tab **Links** > Settings > **Generate link key**, then **Copy key**.
2. Select P&L!B4:E9, click **Export as table**: a green toast confirms the export.
3. PowerPoint: open a blank deck, ribbon tab **pls,fix** > **Links**; tab **Settings**, paste
   the key, **Save**. Tab **Inbox** lists the export; click **Insert**: an editable table lands
   on the slide.
4. Excel: change P&L!C4 to 1000, click **Push all** in the Links tab. PowerPoint: tab
   **Links** > **Update all**: the table repaints in place with the new number. **Break link**
   removes the tracker and leaves the table.

## PowerPoint: object tools

Select two or more shapes on a slide, then ribbon **pls,fix** > **Align left** or **Match
size**: the shapes move or resize; every button in the Objects and Arrange groups acts on the
selection the same way.

## Notes for the reviewer

- The relay at https://dbautomatizacijas.com/modelis/ stores encrypted blobs only; links
  expire after 30 days. No credentials are needed anywhere.
- On Excel for the web, custom functions occasionally show #NAME? on first load while the
  custom-functions runtime starts; reload the workbook once.
- Support: https://dbautomatizacijas.com/modelis/support.html. Source and issues:
  https://github.com/bendiksdaniels/plsfix
