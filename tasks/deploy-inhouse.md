# In-house deployment (Microsoft 365 centralized deployment)

The private launch: an administrator uploads the manifest once in the Microsoft 365 admin
center and the add-in appears in Excel and PowerPoint on Windows, Mac and the web for the
assigned users. No store, no installer on any machine. Written 21.09.2026 at v2.8.20; the
tenant, the group and the people stay out of this public file. Tick as it happens.

## 1. The package (what the administrator gets)

- [x] Manifest: `manifest.prod.xml` of the release to deploy, identical in three places (the
      repo at the tag, the GitHub release asset, the served
      `https://dbautomatizacijas.com/modelis/manifest.xml`); proven for v2.8.20 by one SHA-256
      over all three (`1b0372bd...d68e`).
- [x] Identity: Id `FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E`, Version `2.8.20.0`, ProviderName
      "Daniels Bendiks", hosts Workbook + Presentation, requirement SharedRuntime 1.1, no
      `WebApplicationInfo` (no SSO, no Graph consent prompt).
- [x] Hosted side: every URL in the manifest answers 200 on the public path without Access
      (icons 16/32/64/80, the ribbon icons, both panes, functions.js + functions.json,
      shortcuts.json, support.html); `npx office-addin-manifest validate -p manifest.prod.xml`
      says "The manifest is valid." (21.09).
- [x] The IT one-pager `docs/IT-vienlapa.md` (LV), refreshed against the code on 21.09: CAGR,
      the relay's real limits, the manifest source, the declared minimum.
- [x] Relay limits: the production unit runs with `MODELIS_TRUSTED_PROXY=cloudflare` (gateway
      `server/systemd/plsfix.service`), so the client is the real address; an office behind one
      NAT address is one client (section 5).

## 2. Admin steps (Integrated apps)

1. admin.microsoft.com > Settings > Integrated apps > Upload custom apps.
2. App type Office Add-in; upload the manifest file from the device, or provide the link
   `https://github.com/bendiksdaniels/plsfix/releases/download/v2.8.20/manifest.prod.xml`.
3. Users: a named pilot first (specific users, or one top-level group; nested groups are not
   supported); widen later by editing the assignment, no re-upload.
4. Deployment method Fixed (the tab sits on the ribbon and the user cannot remove it);
   Available and Optional are known to drop the tab from the ribbon at times.
5. Accept (no permissions beyond the add-in itself) > Deploy. Propagation 24-72 h; a running
   Office picks it up at its next launch.

## 3. Proof after propagation (Daniel, 10 min)

- [ ] Excel desktop, fresh launch: the pls,fix tab shows, the pane footer says 2.8.20,
      Ctrl+Shift+M opens the pane, Home > Add-ins lists pls,fix as admin-managed.
- [ ] Excel on the web (office.com): the tab and the pane; `=PLSFIX.CAGR(100,161.051,5)` gives
      0.1 (`#NAME?` while the web functions runtime warms up is documented).
- [ ] PowerPoint desktop and web: the tab, the pane, the Inbox.
- [ ] One round trip: Export selection in Excel, Paste latest linked in PowerPoint; the
      `/version` relay counters move.
- [ ] A colleague's machine: the same checks, no sideload of any kind.

## 4. Afterwards

- Pane and server updates (`deploy.sh modelis`): nothing for the admin. A running desktop shows
  the new pane only after Office quits (shared runtime); Office caches `functions.js` and
  `functions.json` for up to 24 h.
- A manifest change (a new ribbon button, a new function, a new URL): Integrated apps >
  pls,fix > Update with the new file; same Id, higher Version (`release.sh` bumps it); takes
  effect at the next launch.
- Rollback: Integrated apps > pls,fix > remove the app or the assignment. The sideloaded wef
  manifest on the Mac keeps working regardless.
- Traps: "Add-in Error: this add-in is no longer available" after a manifest swap means close
  the pane and load it again from Add-ins; a per-user Upload My Add-in on the web lives in that
  browser only.

## 5. Relay limits for an office behind one address

The relay rates by client. With `MODELIS_TRUSTED_PROXY=cloudflare` the client is the real
address from `cf-connecting-ip`; without it, everything arriving through the gateway shares
ONE bucket (the server prints "behind a proxy and trusting nothing" at start). An office
behind one NAT address is one client either way: 100 writes and 1 200 reads a minute, 64 MiB
of writes a minute, shared by everyone in that office. Raise `MODELIS_RATE_WRITE_PER_MIN` in
the unit when a whole department pushes at once.
