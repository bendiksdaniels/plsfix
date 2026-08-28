#!/bin/sh
# Patch bump per reviewed merge: package.json is the source; the server crate, the
# lockfiles and both generated manifests follow; version:check proves they agree.
set -eu
cd "$(dirname "$0")/.."
npm version patch --no-git-tag-version >/dev/null
V="$(node -p "require('./package.json').version")"
sed -i '' "s/^version = \".*\"/version = \"$V\"/" server/Cargo.toml
cargo check --manifest-path server/Cargo.toml -q
npm run manifest:build >/dev/null
npm run version:check
