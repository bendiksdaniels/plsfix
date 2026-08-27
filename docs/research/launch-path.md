# Launch path: hosting, deployment, AppSource (research agent, 2026-08-27)

Full source URLs in the agent transcript; all claims below are doc-sourced unless flagged.

## Recommendation

Host `dist/` on the existing Hetzner + Cloudflare suite host (one-hostname path), with
**NO Cloudflare Access in front of the add-in path**. Restrict availability at the
Microsoft 365 layer instead (centralized deployment, group assignment). Internal rollout
skips AppSource entirely: admin uploads manifest via M365 admin center > Integrated Apps >
Add-ins, assigns to a group, ~1-3 days propagation. Same manifest GUID across dev/prod;
only SourceLocation/icon/AppDomains/ExtendedOverrides URLs differ; bump Version per change.
Public AppSource launch later: XML manifest still fully accepted (unified NOT required and
NOT supported on perpetual Win/Outlook Mac/mobile anyway), 4-6 weeks with resubmissions.

## Why no auth wall on the pane (confirmed, 4 sources)

- Office on the web renders the pane in an IFRAME; identity providers refuse sign-in pages
  in iframes (Microsoft's own auth overview treats this as the load-bearing fact).
- Desktop just fetches SourceLocation: a 302-to-login turns the pane INTO the login page
  with no way back.
- WKWebView (Mac) has ITP permanently on: third-party cookies blocked, no opt-out.
- Documented failures: WebView2 401 challenge -> "This add-in could not be started"
  (office-js#1436); MSAL loginRedirect breaks once deployed in a pane (msal#5898).
- Any future backend API the pane fetches must also not sit behind cookie-gated Access —
  use token auth or an Access bypass for that path.

## Hosting facts

- HTTPS required (web + AppSource hard rule). No CORS needed to serve a plain task pane.
- office.js MUST load from Microsoft's CDN (Marketplace requirement) — we already do.
- AppDomains = one-way trust list for navigation; subdomains listed separately.
- Two cache layers: Office's own web/Wef cache (clear via Ctrl+F5 in pane / clear-cache
  CLI; unreliable auto-clear) AND normal HTTP caching — serve hashed assets immutable,
  taskpane.html + manifest no-cache. GitHub Pages viable but: ToS discourages commercial
  SaaS, no SLA, no custom Cache-Control (unconfirmed), zero Microsoft support surface.
  The suite host is the better fit.

## Centralized deployment (in-house)

- M365 admin center > Settings > Integrated apps > Add-ins > Deploy > Upload Custom Apps
  (manifest file or URL) > assign Everyone/users/groups (top-level groups only, nested
  unsupported) > Deploy. Needs Exchange admin role. 24-72h to appear.
- Hosted JS/HTML/CSS changes: ZERO admin action. manifest.xml changes: re-upload + Update
  button, takes effect next app launch.
- ExtendedOverrides keyboard shortcuts: no documented restriction under centralized
  deployment (inferred fine — it's just another hosted resource; unconfirmed).
- Current known issues page lists optional/available add-ins sometimes vanishing from the
  ribbon (workaround: deploy as Fixed) — check if rollout misbehaves.

## AppSource (public, later)

- Partner Center > Marketplace offers > Microsoft 365 and Copilot > new offer; publisher
  name must match manifest ProviderName and is immutable. Registration fee: $0 (2026).
- Listing must be FREE to download; monetization only via a connected transactable SaaS
  offer (Fulfillment APIs + licensing DB — a real project) or fully BYO licensing.
- Top rejection causes: support URL (https page, not email/GitHub/doc), privacy policy URL
  (specific to the app, separate from ToU), EULA, missing test notes ("automatically
  fail"), icon/schema/version-mismatch issues.
- Policy 1120.3: must work on web + Mac for every API in Requirements; task-pane add-ins
  must have add-in commands (we do); must work on touch-only devices — keyboard shortcuts
  must not be the only path to any core function (they aren't; every action has a button).
- Review: 3-4 business days per round; 4-6 weeks typical total. Updates: same Id, bumped
  Version, users get an in-app Update prompt (Excel/PPT/Word).

## C8 actions derived

1. manifest.prod.xml: same GUID, prod URLs (suite host path), Version 1.0.0.0-aligned.
2. Decide the suite path key with Daniel (gateway integration = suite-repo change with its
   own deploy rules; Cloudflare Access must EXCLUDE this path — needs his Access config).
3. vite build already hashes assets; ensure taskpane.html served no-cache on the host.
4. For AppSource later: draft privacy policy + support page + screenshots + test notes.
