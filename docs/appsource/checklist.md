# AppSource submission checklist for pls,fix

The store is the real "install and use" for an Office add-in: Excel > Add-ins > search
"pls,fix" > Add, on Windows, Mac and the web, auto-updated, no manifest. Free listing.
Sources: Microsoft's [publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist)
and [step-by-step submission guide](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide)
(read 12.09.2026). Expect 4-6 weeks end to end; first submissions usually bounce once;
3-4 business days per review round.

## Before the account (repo side)

- [ ] Store validation green: `npx office-addin-manifest validate -p manifest.prod.xml`.
      12.09 status: two errors left. "High Resolution Icon Missing" (a 64x64
      `HighResolutionIconUrl`) lands with the feature wave; "Icon URL Unreachable" was the
      suite's country gate answering 403 to `MicrosoftOfficeStoreValidationService`
      (VPS nginx log 12.09), fixed in the gateway by exempting `/modelis/`.
- [ ] `SupportUrl` in `manifest/spec.ts` = https://dbautomatizacijas.com/modelis/support.html
      (an https page, not an e-mail, not the GitHub repo; feature wave K3).
- [ ] `helpUrl` on both custom functions (`src/functions/metadata.ts`; feature wave M2).
- [ ] `public/privacy.html` served and identical to `docs/appsource/privacy-policy.md`
      (names the app, the relay, personal data handling; a Terms of Use page does not count).
- [ ] `public/support.html` served (how to reach us; issues link).
- [ ] Screenshots, at least one at 1366 x 768 (`docs/appsource/screenshots/`).
- [ ] `ProviderName` stays "Daniels Bendiks" (it must equal the Partner Center publisher name
      and cannot change after the offer is created).
- [ ] Policy 1120.3: every feature works on the web and on Mac, and every action has a button
      (keyboard shortcuts are never the only path). True today; re-check after the wave.

## The account (Daniel)

- [ ] Partner Center: open a developer account and join the **Microsoft 365 and Copilot**
      program (work account recommended; legal name; sign the publisher agreement;
      registration is free). Publisher display name = "Daniels Bendiks".
- [ ] Partner Center > Marketplace offers > Microsoft 365 and Copilot > **+ New offer** >
      Office Add-in; name "pls,fix"; check availability; associate the publisher.

## The offer (paste from `listing.md`)

- [ ] Product setup: no Apple Store, no Entra ID, no additional purchases, no CRM.
- [ ] Packages: upload `manifest.prod.xml` (from the latest release); wait for "Complete".
- [ ] Properties: categories (up to 3), industries (up to 2), Standard Contract EULA,
      privacy link, support link.
- [ ] Marketplace listings: English; summary, description, keywords, icons, screenshots.
- [ ] Availability: all markets, publish on approval, free.
- [ ] Notes for certification: paste `test-notes.md` (includes the custom-function test the
      store requires; reviewers cannot contact you for missing information).
- [ ] Review and publish > Submit.

## After approval

- [ ] Updates: same add-in Id, bumped Version in the manifest, resubmit; users get an in-app
      update prompt. Pane-only changes need no resubmission (the store points at our URLs).
- [ ] Watch the relay: a public listing brings strangers' traffic; the server rate-limits per
      IP and caps storage at 1 GiB, and the Cloudflare edge rule is the cheap extra layer.
