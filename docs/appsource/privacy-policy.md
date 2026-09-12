# pls,fix privacy policy

Effective 12 September 2026. This policy covers the pls,fix add-in for Microsoft Excel and
PowerPoint and the pls,fix link relay at `https://dbautomatizacijas.com/modelis/`, both
published by Daniels Bendiks ("we"). It is the policy the Microsoft AppSource listing links to;
the served copy lives at `https://dbautomatizacijas.com/modelis/privacy.html`.

## What pls,fix is

pls,fix is a task-pane add-in that formats, audits and links financial models. It runs inside
Excel and PowerPoint and reads the workbook or deck that is open in the application. There is
no account, no sign-in, no telemetry, no analytics and no advertising. The add-in code itself
is loaded from our server, and the Office.js library from Microsoft's content delivery
network, the same way every Office web add-in loads.

## Data on your device

- Settings you choose (the brand palette, the three paintbrush slots, the language, the
  "seen the welcome card" flag) are stored inside the workbook you are working on and in the
  Office application's local storage on your device. They never leave the device unless you
  share the workbook.
- Everything the tools do (formatting, auditing, model checks, charts, templates) happens
  inside the application. Nothing is uploaded.

## Data that leaves your device: linked objects

The one feature that sends data anywhere is **Links** (Excel -> PowerPoint linked objects),
and only when you use it:

- When you export a range, table, text or chart, the add-in renders it, encrypts it on your
  device (AES-GCM; the keys are derived from a link key that exists only inside your workbook
  and your deck) and uploads the encrypted blob to the relay. The relay stores the encrypted
  blob, a one-way hash of the link's authorisation key, the blob's size and timestamps. It
  cannot read the content, does not know the workbook's name and holds no user identity.
- PowerPoint downloads the encrypted blob and decrypts it on your device.
- A link is deleted 30 days after its last push, an item waiting in the PowerPoint inbox after
  7 days; "Break link" or "Remove" deletes it at once. Anyone holding a deck that contains a
  link key can fetch that link's encrypted blob for those 30 days, so break links before a
  deck leaves your organisation.

## Server logs

The server that hosts the add-in and the relay (Hetzner Online GmbH, Helsinki, Finland,
reached through Cloudflare's network) keeps standard web-server access logs: IP address,
country, time, requested path, response code and browser identification. They serve
security, abuse prevention and capacity planning, are kept for 14 days and are then deleted.
They are not combined with anything else.

## Sharing

We do not sell, rent or share data with anyone. The only parties that process traffic are the
hosting provider (Hetzner) and the network provider (Cloudflare), each under their own terms.
No data is transferred for marketing.

## Your rights

Under the GDPR you may ask what we hold about you and ask for its deletion. Given the design
above, the answer is normally "nothing beyond a 14-day log line"; encrypted relay blobs cannot
be attributed to a person. Requests go through the support page linked from the add-in.

## Changes

Changes to this policy are published on the same page with a new effective date.
