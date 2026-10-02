#!/bin/sh
# One reviewed merge = one version: bump (scripts/bump-patch.sh), commit "vX.Y.Z"
# and write the annotated tag in a single step, so no version lands untagged again.
# Refuses a dirty tree: the bump commit carries the bump and nothing else.
set -eu
cd "$(dirname "$0")/.."
LEVEL="${1:-patch}"
case "$LEVEL" in patch|minor|major) ;; *) echo "usage: sh scripts/release.sh [patch|minor|major]"; exit 2;; esac
if [ -n "$(git status --porcelain)" ]; then
  echo "release: commit or stash your changes first (the bump commit must be only the bump)"; exit 1
fi
sh scripts/bump-patch.sh "$LEVEL"
V="$(node -p "require('./package.json').version")"
git add package.json package-lock.json server/Cargo.toml server/Cargo.lock manifest.xml manifest.prod.xml public/shortcuts.html
git commit -q -m "v$V"
git tag -a "v$V" -m "v$V"
echo "released v$V ($(git rev-parse --short HEAD)), tagged"
