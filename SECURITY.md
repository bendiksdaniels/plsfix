# Security

## What leaves the machine

The panes run inside Office and read the workbook in the Excel process. A workbook with no
linked objects sends nothing anywhere; the only network traffic is loading the pane itself.

A linked object (a range picture, a table, a text) is encrypted in the Excel pane before it is
uploaded: AES-GCM with keys derived (HKDF) from a per-link token that exists only in the
workbook's registry and in the deck's shape tags. The relay stores the ciphertext and
`sha256(authKey)`, never plaintext, file names, workbook names or user identities. PowerPoint
fetches the ciphertext and decrypts it in its own pane. A link expires 30 days after its last
push, an inbox item after 7 days; a deck that holds a link key can pull that link for that
long, so break links before a deck leaves the organisation.

## The host

- The pane URL cannot sit behind a login wall (Office webviews cannot complete one), so the
  hosted path is open and serves built assets only; no dotfile is ever served, with the
  percent-encoded forms folded first.
- Writes are rate-limited per client IP (429 with `Retry-After`), the store has a byte ceiling
  (507) and a per-workspace inbox cap; expired rows are swept hourly.
- `/manifest.xml` is the one document served on purpose: it is the file users sideload.

## Reporting a vulnerability

Use GitHub's private vulnerability reporting on this repository (Security > Report a
vulnerability). Please keep the details private until a fix is out; expect an answer within a
week.
