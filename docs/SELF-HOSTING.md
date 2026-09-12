# Self-hosting pls,fix

Run your own pls,fix host: the panes, the custom functions and the encrypted link relay, served
from an HTTPS address you control. Ten minutes with Docker; the bare binary works too.

## What you need

- A Linux machine (any Docker host) reachable at an HTTPS address, say
  `https://models.example.com/`. Office loads add-ins over HTTPS only; plain HTTP is accepted
  for `localhost` alone.
- Docker with Compose, or Node.js 22 plus a recent stable Rust for the bare build.
- A DNS name for the machine if Caddy is to fetch the certificate for you.

## 1. Docker Compose (recommended)

```bash
git clone https://github.com/bendiksdaniels/plsfix.git
cd plsfix/deploy
cp .env.example .env                    # set PLSFIX_PUBLIC_URL and PLSFIX_DOMAIN
docker compose --profile https up -d
```

That pulls `ghcr.io/bendiksdaniels/plsfix:latest`, starts it on `127.0.0.1:8804`, and starts
Caddy in front of it; Caddy fetches a Let's Encrypt certificate for `PLSFIX_DOMAIN` and proxies
port 443 to the add-in. Already running nginx, Traefik or your own Caddy? Leave the profile out
(`docker compose up -d`) and proxy your HTTPS host to `127.0.0.1:8804`.

Check: `curl https://models.example.com/healthz` answers `{"ok":true}` and
`curl https://models.example.com/version` names the version.

## 2. Hand out the manifest

`https://models.example.com/manifest.xml` is the add-in manifest for YOUR host. The server takes
the shipped `manifest.prod.xml` and re-points it: every URL at `PLSFIX_PUBLIC_URL`, the trusted
AppDomain at your origin, and the add-in id derived from your URL (UUID v5), so your copy and the
hosted pls,fix can sit side by side in one Excel. Download it and sideload it exactly like the
hosted manifest: [INSTALL.md](INSTALL.md), path 1, with your file instead.

To pin a version instead of `latest`, set `image: ghcr.io/bendiksdaniels/plsfix:v2.6.16` in
`docker-compose.yml`; the [releases page](https://github.com/bendiksdaniels/plsfix/releases)
lists the tags, and every release carries the same image.

## 3. Updating

```bash
docker compose pull && docker compose --profile https up -d
```

Users get the new panes on their next Office launch; the manifest never changes for a pane-only
update. Office caches the custom functions file separately, so `=PLSFIX.ROUND` can lag by up to
a day.

## Environment variables

| Variable                     | Default             | Meaning                                                                                                              |
| ---------------------------- | ------------------- | -------------------------------------------------------------------------------------------------------------------- |
| `MODELIS_PUBLIC_URL`         | unset               | The public base URL. Set: `/manifest.xml` is re-pointed at it. Unset: the shipped manifest verbatim (the hosted add-in). |
| `MODELIS_BIND`               | `127.0.0.1`         | Listen address. The image sets `0.0.0.0`.                                                                            |
| `MODELIS_PORT`               | `8804`              | Listen port.                                                                                                         |
| `MODELIS_STATIC`             | `dist`              | The built panes. The image: `/app/static`.                                                                           |
| `MODELIS_DATA`               | `data`              | Where `relay.sqlite` lives. The image: `/data`, a volume.                                                            |
| `MODELIS_MANIFEST`           | `manifest.prod.xml` | The file `/manifest.xml` serves. The image: `/app/manifest.prod.xml`.                                                |
| `MODELIS_MAX_BYTES`          | 1 GiB               | Relay storage ceiling; a push past it answers 507.                                                                   |
| `MODELIS_RATE_WRITE_PER_MIN` | 300                 | Relay writes per minute per client IP (429 with `Retry-After` past it).                                              |
| `MODELIS_RATE_READ_PER_MIN`  | 1200                | Relay reads per minute per client IP.                                                                                |
| `MODELIS_TRUSTED_PROXY`      | unset               | Which forwarding header names the client for the rate limits: unset or `none` = the peer address, `cloudflare` = `CF-Connecting-IP`, `xff` = the last `X-Forwarded-For` hop. |

The `MODELIS_` prefix and the binary name `plsfix-server` are historical (the first host was
the `modelis` key of a tools suite) and stay, so existing deployments keep working.

## What the relay stores

Linked pictures, tables and texts travel Excel -> relay -> PowerPoint encrypted in the pane
(AES-GCM, keys derived from a per-link token that lives only in the workbook and the deck's
shape tags). The server holds the ciphertext and `sha256(authKey)` per link: no plaintext, no
file names, no user identities. Links expire 30 days after their last push, inbox items after
7 days, swept hourly. A backup is a copy of the `plsfix-data` volume; nothing else holds state.

Rate limits and the byte ceiling protect a host that is open to the internet. Put your proxy's
own rate limiting in front as well; the pane URL itself cannot sit behind a login page
(Office webviews cannot complete one, see `docs/research/launch-path.md`), so restrict who
receives the manifest, not the URL.

### Tell the server which proxy to trust

The rate limits are per client, so the server has to know which client a request came from.
Behind a proxy every connection arrives from the proxy, and the real address is in a header
that the client can also write, so nothing is trusted unless you say so: with
`MODELIS_TRUSTED_PROXY` unset, every request is counted under the peer address. Behind the
Caddy in `docker-compose.yml`, behind nginx, or behind any proxy that appends to
`X-Forwarded-For`, set `MODELIS_TRUSTED_PROXY=xff`: the server takes the LAST hop, the one
your proxy wrote, never the first, which the client chose. Behind Cloudflare, set
`MODELIS_TRUSTED_PROXY=cloudflare` for `CF-Connecting-IP`. Set it only when the proxy really
does rewrite that header on every request and the origin cannot be reached around it -
otherwise a client picks its own bucket by sending its own header, and the limits stop
limiting anything. Leaving it unset behind a proxy is safe but blunt: every client shares one
bucket, so a busy team will see 429s.

## The bare binary

```bash
npm ci && npm run build                                               # -> dist/
cargo build --release --locked --manifest-path server/Cargo.toml     # -> server/target/release/plsfix-server
MODELIS_BIND=0.0.0.0 MODELIS_STATIC=dist MODELIS_DATA=/var/lib/plsfix \
MODELIS_PUBLIC_URL=https://models.example.com/ ./server/target/release/plsfix-server
```

`deploy/plsfix.service` is a systemd unit to copy from (adjust the paths, add
`MODELIS_PUBLIC_URL`). HTTPS still comes from your reverse proxy.
