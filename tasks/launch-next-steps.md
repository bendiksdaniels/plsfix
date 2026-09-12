# pls,fix launch: next steps (paused 13.09.2026, 00:1x)

State at the pause: **v2.7.1 LIVE**, repo PUBLIC under MIT, one-line installers on every release,
AppSource kit written. Everything below is what stands between today and "Excel > Add-ins >
search pls,fix > Add". Owner in brackets. Resume with this file, then `tasks/AUTORESUME.md`.

## 1. Prove the install on real machines [Daniel, 15 min]

- [ ] Windows: in PowerShell, `irm https://github.com/bendiksdaniels/plsfix/releases/latest/download/plsfix-install-windows.ps1 | iex`,
      reopen Excel and PowerPoint, the pls,fix tab shows, Ctrl+Shift+M opens the pane. Then the
      uninstall line (same URL with `uninstall`) and a relaunch: the tab is gone. The script was
      parsed and dry-run on PowerShell 7 here; the registry write itself has run only on paper.
- [ ] Mac: the `curl ... | sh` line from the README on a Mac that never had pls,fix (the
      script was proven against a fake home folder; your own Mac still carries the dev-era
      `manifest.prod.xml` in its wef folders, which the installer leaves alone).
- [ ] `tasks/launch-check.md`: the desktop pass of the tools themselves (25 min, Mac first,
      Windows once). A failing action's toast "Copy details" is the whole bug report.

## 2. Close the store validator [feature wave, then Daniel to confirm]

- [ ] `npx office-addin-manifest validate -p manifest.prod.xml` must print "The manifest is
      valid." Today: one error left, the 64 px `HighResolutionIconUrl` (wave slice K3 adds it).
      The other error, "Icon URL Unreachable", was the gateway's Latvia-only country gate and is
      fixed and deployed (the hosting gateway 056d9c3).
- [ ] After K3 lands: SupportUrl = `https://dbautomatizacijas.com/modelis/support.html` and
      `public/privacy.html` served; copy `docs/appsource/privacy-policy.md` into that page (or
      the other way round) so the store link and the repo text are one source.
- [ ] `helpUrl` on both custom functions (wave slice M2).
- [ ] Re-upload nothing yet: the M365 in-house centralized upload takes the manifest of
      whichever version passes the validator (SupportUrl changed, so it is a new manifest).

## 3. Screenshots for the listing [Daniel or the web rig, 20 min]

At least one at 1366 x 768; three planned, captions in `docs/appsource/listing.md`:
1. Excel pane on the demo model: Autocolor + audit overlay on the P&L.
2. PowerPoint: a table and a chart inserted from the model, Links tab open.
3. Excel Tools tab: format cycles, paintbrush slots, the six calculation blocks.
Save to `docs/appsource/screenshots/`. Easiest: Excel on the web with the demo workbook from
the release, browser window sized to 1366 x 768, system screenshot. The rig
(`scripts/rig/README.md`) can do it once the scratch Chrome is signed in.

## 4. Partner Center [Daniel, 30 min of forms, then 4-6 weeks of waiting]

- [ ] Open a Partner Center developer account, join the **Microsoft 365 and Copilot** program
      (registration free; work account recommended; legal name; sign the publisher agreement).
      Publisher display name **"Daniels Bendiks"**, identical to the manifest's ProviderName;
      it cannot change after the offer exists.
- [ ] Marketplace offers > Microsoft 365 and Copilot > **+ New offer** > Office Add-in > name
      "pls,fix" > Check availability > Create.
- [ ] Fill the forms from `docs/appsource/listing.md` (product setup, package =
      `manifest.prod.xml` of the validated version, properties, listing text, keywords,
      icons, screenshots, availability all markets, free) and paste
      `docs/appsource/test-notes.md` into **Notes for certification** (it includes the
      custom-function test the store demands; reviewers cannot contact you for missing info).
- [ ] Review and publish > Submit. Expect 3-4 business days per round, one bounce is normal,
      4-6 weeks total. Fix what the rejection names, resubmit; same add-in Id, bumped Version.
- [ ] `docs/appsource/checklist.md` is the tick list for all of the above.

## 5. Hosting hygiene before strangers arrive [Daniel + a session]

- [ ] Cloudflare rate-limit rule on `/modelis/api/*` (dashboard, free plan: one rule). The
      server already limits per IP (300 writes / 1200 reads per minute) and caps the relay at
      1 GiB; the edge rule stops a flood before it reaches the VPS.
- [ ] The feature wave's relay hardening (`MODELIS_TRUSTED_PROXY`, per-client byte budget):
      when it lands, add the paragraph to `docs/SELF-HOSTING.md` and set the unit's env var at
      deploy (the wave's session owns it).
- [ ] Decide on the `check` workflow: disabled by Daniel ("it keeps saying something failed");
      the flake behind every failure is fixed as of v2.7.1 (`test/hung-sync.ts`, no fake-clock
      copies remain). `gh workflow enable check` turns it back on; the release workflow is
      unaffected and green.
- [ ] Watch `/version` on the live host: `relay.links` and `relay.bytes` show public use.

## 6. After approval

- [ ] README "Get pls,fix": the AppSource line becomes the first line ("Excel > Add-ins >
      search pls,fix"), the installers stay as the second path for locked-down tenants.
- [ ] Manual (`npm run manual`) regenerated for the shipped version, as a separate commit.
- [ ] `ROADMAP.md` tick-through (Daniel's file; deltas listed in `docs/FEATURES.md`).

## Where things are

- Installers: `deploy/install/` (attached to every release by `.github/workflows/release.yml`).
- Store kit: `docs/appsource/` (privacy-policy, listing, test-notes, checklist).
- Install and self-host guides: `docs/INSTALL.md`, `docs/SELF-HOSTING.md`.
- Gateway change for the world-open add-in path: `the hosting gateway` `nginx/gen_servers.py`
  `OPEN_COUNTRY_PATHS`.
- Ledger with the full history of 12.09: `tasks/AUTORESUME.md`.
