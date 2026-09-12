# Get pls,fix

pls,fix is free (MIT). Three ways to get it; the first needs no setup at all.

| You want to                                       | Path                                         |
| ------------------------------------------------- | -------------------------------------------- |
| Use it in your own Excel and PowerPoint today     | **1. The hosted add-in** (this page)         |
| Run it on your own server for your team           | **2. Self-host** ([SELF-HOSTING.md](SELF-HOSTING.md)) |
| Change it or run it from source                   | **3. Build it** ([CONTRIBUTING.md](../CONTRIBUTING.md)) |

Requirements: Excel and PowerPoint with Microsoft 365 (or Excel 2019 and PowerPoint 2021 or
later), on Windows, Mac or the web. Nothing is installed on the computer: the manifest file
below only tells Office where to load the panes from.

## 1. The hosted add-in

Every release carries `manifest.prod.xml`, the one file to sideload. Download it:

- <https://github.com/bendiksdaniels/plsfix/releases/latest/download/manifest.prod.xml>
- or <https://dbautomatizacijas.com/modelis/manifest.xml>, the same file from the host itself.

Sideload it into Excel, and again into PowerPoint if you want linked objects and the slide
tools. The steps come from Microsoft's own pages, linked under each platform.

### Mac

1. In Finder press Cmd+Shift+G and open
   `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (create the `wef` folder if
   it does not exist). For PowerPoint the folder is
   `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`.
2. Copy `manifest.prod.xml` into the folder.
3. Quit and reopen Excel (or PowerPoint), open a workbook, then **Home > Add-ins** and pick
   **pls,fix**. The **pls,fix** ribbon tab appears; Ctrl+Shift+M opens the pane.

Source: [Sideload Office Add-ins on Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac).

### Office on the web

1. Open a workbook (or a deck) at <https://office.com>.
2. **Home > Add-ins > More Settings**, then **Upload My Add-in**.
3. **Browse** to `manifest.prod.xml` and **Upload**.

The add-in is remembered in that browser's local storage, so a cleared cache or another
browser means uploading it again. On the web, `=PLSFIX.ROUND` and `=PLSFIX.ROUNDSUM` can
show `#NAME?` when Office's custom-functions runtime fails to start; every button still works.

Source: [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

### Windows desktop

Windows has no "upload a file" button; it reads manifests from a shared folder you trust once.

1. Put `manifest.prod.xml` in a folder and share it: right-click the folder > **Properties** >
   **Sharing** > **Share**, add yourself, note the network path (`\\YOUR-PC\plsfix`).
2. In Excel: **File > Options > Trust Center > Trust Center Settings > Trusted Add-in
   Catalogs**. Enter the network path as **Catalog Url**, **Add catalog**, tick **Show in
   Menu**, **OK** twice, then close and reopen Excel.
3. **Home > Add-ins > Advanced**, choose **SHARED FOLDER** at the top of the dialog, select
   **pls,fix**, **Add**. Repeat step 3 in PowerPoint (the catalog is shared by all Office apps).

Microsoft calls this a testing route: a release that changes the ribbon (a new button) means
adding the add-in again from the shared folder. For a whole company the admin route below is
the proper one.

Source: [Sideload Office Add-ins on Windows from a network share](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

### A whole organisation (Microsoft 365 admin)

An administrator deploys it once for everyone: Microsoft 365 admin center > **Settings >
Integrated apps > Upload custom apps**, upload `manifest.prod.xml`, assign users or top-level
groups, **Deploy**. The add-in shows up within a day or three; pane updates need no admin
action afterwards, a manifest change (a new ribbon button) needs a re-upload. Needs a Global
or Exchange administrator.

Sources: [Integrated apps portal](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps),
[Deploy add-ins in the admin center](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins).

### First steps

Download `pls,fix Demo Model.xlsx` from the same release and open it: the **Start here**
sheet is a checklist that walks through every tool on sheets built for them (a P&L with two
planted errors, a bridge table for the waterfall, a tornado block, a rounding split, a wide
grid to unpivot). Every section of both panes carries a `?` that explains its buttons, and the
Tools tab prints the keyboard shortcut card.

### What leaves your computer

The panes read the workbook inside Excel. A workbook with no linked objects sends nothing
anywhere. A linked object (a range picture, a table or a text sent to PowerPoint) is
encrypted in the pane before it is uploaded; the relay at `dbautomatizacijas.com/modelis/`
stores ciphertext only, for 30 days, and deletes it after. Details: [SECURITY.md](../SECURITY.md).

### Removing it

Delete the manifest from the `wef` folder (Mac) or the shared folder (Windows), then clear
the Office cache: [Clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache).
On the web, remove it under **Home > Add-ins > More Settings** or clear the browser's storage.

## 2. Self-host

Your own host, your own relay, the same panes: [SELF-HOSTING.md](SELF-HOSTING.md). Ten
minutes with `docker compose`; the server hands out a manifest re-pointed at your address.

## 3. Build it

`git clone`, `npm install`, `npm start` sideloads the development manifest into desktop
Excel with the panes served from `https://localhost:3000`: [CONTRIBUTING.md](../CONTRIBUTING.md).
