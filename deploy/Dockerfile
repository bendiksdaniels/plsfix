# syntax=docker/dockerfile:1
# pls,fix in one image: the built panes plus the Rust host and link relay.
# Build from the repo root: docker build -f deploy/Dockerfile -t plsfix .
# Run with MODELIS_PUBLIC_URL=https://your.host/ so /manifest.xml points at
# this deployment; the encrypted relay lives in /data (mount a volume).

FROM node:22-bookworm-slim AS pane
WORKDIR /src
COPY package.json package-lock.json ./
RUN npm ci --no-audit --no-fund
COPY . .
RUN npm run build

FROM rust:1.98-bookworm AS host
ARG PLSFIX_VERSION=dev
ENV PLSFIX_VERSION=$PLSFIX_VERSION
WORKDIR /src/server
COPY server/ ./
RUN cargo build --release --locked

FROM debian:bookworm-slim
LABEL org.opencontainers.image.source="https://github.com/bendiksdaniels/plsfix" \
      org.opencontainers.image.licenses="MIT" \
      org.opencontainers.image.description="pls,fix: Excel and PowerPoint add-in host and link relay"
RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl \
    && rm -rf /var/lib/apt/lists/* \
    && useradd --system --uid 10001 --create-home plsfix \
    && mkdir -p /data && chown plsfix:plsfix /data
WORKDIR /app
COPY --from=host /src/server/target/release/plsfix-server /app/bin/plsfix-server
COPY --from=pane /src/dist /app/static
COPY manifest.prod.xml /app/manifest.prod.xml
ENV MODELIS_BIND=0.0.0.0 \
    MODELIS_PORT=8804 \
    MODELIS_STATIC=/app/static \
    MODELIS_MANIFEST=/app/manifest.prod.xml \
    MODELIS_DATA=/data
USER plsfix
VOLUME /data
EXPOSE 8804
HEALTHCHECK --interval=30s --timeout=3s CMD curl -fsS http://127.0.0.1:8804/healthz || exit 1
ENTRYPOINT ["/app/bin/plsfix-server"]
