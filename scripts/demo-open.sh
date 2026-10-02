#!/bin/sh
# Build the demo workbook, put the production manifest into every Office wef
# folder, then open Excel on the workbook and PowerPoint on a fresh deck.
# Mac only; Office reads the manifest at launch, so an app that was already
# running keeps whatever it loaded until it is quit and reopened.
set -eu
[ "$(uname)" = Darwin ] || { echo "demo-open: Mac only" >&2; exit 1; }
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WORKBOOK="$ROOT/demo/out/pls,fix Demo Model.xlsx"
cargo run --quiet --manifest-path "$ROOT/demo/Cargo.toml"
for app in com.microsoft.Excel com.microsoft.Powerpoint; do
  mkdir -p "$HOME/Library/Containers/$app/Data/Documents/wef"
done
sh "$ROOT/scripts/wef-restore-prod.sh"
open -a "Microsoft Excel" "$WORKBOOK"
osascript -e 'tell application "Microsoft PowerPoint"' -e 'activate' -e 'make new presentation' -e 'end tell' >/dev/null 2>&1 \
  || open -a "Microsoft PowerPoint"
echo "demo: Excel opened on $(basename "$WORKBOOK"), PowerPoint on a new deck, manifest v$(node -p "require('$ROOT/package.json').version")"
echo "demo: no pls,fix tab? quit that app and run npm run demo again (manifests load at app start)"
