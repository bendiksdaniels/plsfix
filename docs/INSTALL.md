# Get pls,fix

pls,fix is free (MIT). Nothing is installed on the computer: one line tells Excel and
PowerPoint where to load the add-in from, and every later update reaches you on the next
launch. The Microsoft AppSource listing (Excel > Add-ins > search "pls,fix") is in
preparation; until it is live, this page is the way in.

Requirements: Excel and PowerPoint with Microsoft 365 (or Excel 2019 and PowerPoint 2021 or
later), on Windows, Mac or the web.

## Windows

Open PowerShell (Start, type "PowerShell", Enter) and paste:

```powershell
irm https://github.com/bendiksdaniels/plsfix/releases/latest/download/plsfix-install-windows.ps1 | iex
```

Close and reopen Excel and PowerPoint. To remove it, run the same line with `uninstall` in
place of `install`. What the line does: downloads the add-in manifest to `%LOCALAPPDATA%\plsfix`
and registers it under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`, the key Microsoft's
own add-in tooling uses for sideloaded add-ins. No admin rights, no other software.

## Mac

Open Terminal (Spotlight, type "Terminal") and paste:

```bash
curl -fsSL https://github.com/bendiksdaniels/plsfix/releases/latest/download/plsfix-install-mac.command | sh
```

Quit and reopen Excel and PowerPoint. To remove it, run the same line with `uninstall` in
place of `install`. What the line does: copies the manifest into
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` and the PowerPoint twin, the
folders Office reads sideloaded add-ins from. The same file can be downloaded and
double-clicked instead; macOS then asks you to allow it under System Settings > Privacy &
Security, which the Terminal line avoids.

## Office on the web

1. Open a workbook (or a deck) at <https://office.com>.
2. **Home > Add-ins > More Settings**, then **Upload My Add-in**.
3. **Browse** to [`manifest.prod.xml`](https://github.com/bendiksdaniels/plsfix/releases/latest/download/manifest.prod.xml)
   and **Upload**.

The add-in is remembered in that browser's local storage, so a cleared cache or another
browser means uploading it again. On the web, `=PLSFIX.ROUND` and `=PLSFIX.ROUNDSUM` can show
`#NAME?` when Office's custom-functions runtime fails to start; every button still works.
Source: [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

## After the install

The **pls,fix** tab sits on the ribbon of both apps; Ctrl+Shift+M opens the Excel pane, or
**Home > Add-ins > pls,fix**. Download `pls,fix Demo Model.xlsx` from the
[latest release](https://github.com/bendiksdaniels/plsfix/releases/latest) and open it: the
**Start here** sheet is a checklist that walks through every tool on sheets built for them (a
P&L with two planted errors, a bridge table for the waterfall, a tornado block, a rounding
split, a wide grid to unpivot). Every section of both panes carries a `?` that explains its
buttons, and the Tools tab prints the keyboard shortcut card.

`pls,fix Demo Deck.pptx` from the same release demos the PowerPoint Slide and Where
pickers: pick a target slide and a spot, then Insert or Update all.

## For a whole organisation (Microsoft 365 admin)

An administrator deploys it once for everyone: Microsoft 365 admin center > **Settings >
Integrated apps > Upload custom apps**, upload `manifest.prod.xml`, assign users or top-level
groups, **Deploy**. The add-in shows up within a day or three; pane updates need no admin
action afterwards, a manifest change (a new ribbon button) needs a re-upload. Needs a Global
or Exchange administrator. Sources: [Integrated apps portal](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps),
[Deploy add-ins in the admin center](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins).

## Manual sideload, if you would rather not run a script

Download [`manifest.prod.xml`](https://github.com/bendiksdaniels/plsfix/releases/latest/download/manifest.prod.xml)
(the same file is served at <https://dbautomatizacijas.com/modelis/manifest.xml>).

- **Mac:** in Finder press Cmd+Shift+G, open
  `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (create `wef` if missing), copy
  the file in; the same for `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`.
  Restart the apps, then **Home > Add-ins > pls,fix**.
  Source: [Sideload Office Add-ins on Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac).
- **Windows:** Office reads manifests from a shared folder you trust once. Put the file in a
  folder and share it (right-click > **Properties > Sharing > Share**, note the network path).
  In Excel: **File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs**,
  enter the path as **Catalog Url**, **Add catalog**, tick **Show in Menu**, restart Excel.
  Then **Home > Add-ins > Advanced > SHARED FOLDER > pls,fix > Add**; the same in PowerPoint.
  Source: [Sideload Office Add-ins on Windows from a network share](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

Removing a manual sideload: delete the file from the folder and clear the Office cache
([Clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache)).

## What leaves your computer

The panes read the workbook inside Excel. A workbook with no linked objects sends nothing
anywhere. A linked object (a range picture, a table or a text sent to PowerPoint) is
encrypted in the pane before it is uploaded; the relay at `dbautomatizacijas.com/modelis/`
stores ciphertext only, for 30 days, and deletes it after. Details: [SECURITY.md](../SECURITY.md).

A fresh install needs no relay for links: by default they move from Excel to PowerPoint by
copy and paste on the same computer, and the relay above is a switch in each pane for anyone
who wants it.

## Your own server, or the source

[SELF-HOSTING.md](SELF-HOSTING.md): `docker compose up` on your machine; its
`https://your.host/manifest.xml` works with the same installers
(`$env:PLSFIX_MANIFEST_URL` on Windows, `PLSFIX_MANIFEST_URL=` or the URL as the argument on
Mac). [CONTRIBUTING.md](../CONTRIBUTING.md): `git clone`, `npm install`, `npm start`.
