#!/bin/sh
# Patch bump per reviewed merge: package.json is the source; the server crate, the
# lockfiles and both generated manifests follow; version:check proves they agree.
set -eu
cd "$(dirname "$0")/.."
LEVEL="${1:-patch}"
npm version "$LEVEL" --no-git-tag-version >/dev/null
V="$(node -p "require('./package.json').version")"
# Only the [package] version: an unanchored sed would also rewrite the version
# of the first dependency written as its own [dependencies.x] table.
PLSFIX_VERSION="$V" node -e '
const fs = require("node:fs");
const file = "server/Cargo.toml";
const lines = fs.readFileSync(file, "utf8").split("\n");
let inPackage = false;
let done = false;
for (let i = 0; i < lines.length; i += 1) {
  if (lines[i].startsWith("[")) inPackage = lines[i].trim() === "[package]";
  else if (inPackage && !done && /^version\s*=/.test(lines[i])) {
    lines[i] = "version = \"" + process.env.PLSFIX_VERSION + "\"";
    done = true;
  }
}
if (!done) {
  process.stderr.write("bump-patch: no [package] version in " + file + "\n");
  process.exit(1);
}
fs.writeFileSync(file, lines.join("\n"));
'
# Lockfile only: a full install can move transitive dependencies within their
# semver ranges, and that churn has no business landing in a version bump.
npm install --package-lock-only --no-audit --no-fund >/dev/null
cargo check --manifest-path server/Cargo.toml -q
npm run manifest:build >/dev/null
# The shortcut card embeds the version, so it is regenerated with the manifests.
npm run shortcuts:build >/dev/null
npm run version:check
