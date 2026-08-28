#!/bin/sh
# After `npm stop` (which unregisters the dev manifest) put the production
# manifest back into every Office wef folder that exists, so a fresh Excel or
# PowerPoint launch loads the add-in from the server again. Mac only.
set -eu
[ "$(uname)" = Darwin ] || exit 0
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
ID="FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E"
for app in com.microsoft.Excel com.microsoft.Powerpoint; do
  wef="$HOME/Library/Containers/$app/Data/Documents/wef"
  [ -d "$wef" ] || continue
  cp "$ROOT/manifest.prod.xml" "$wef/$ID.manifest.xml"
  echo "restored prod manifest for $app"
done
