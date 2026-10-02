#!/bin/sh
# Installs pls,fix into Excel and PowerPoint on this Mac: downloads the add-in
# manifest and puts it where Office reads sideloaded add-ins. No admin rights,
# nothing else installed; the panes load from the pls,fix server.
# Usage: curl -fsSL <this file's URL> | sh        (or double-click it)
#        sh plsfix-install-mac.command https://your.host/manifest.xml   (self-hosted)
set -eu

URL="${1:-${PLSFIX_MANIFEST_URL:-https://github.com/bendiksdaniels/plsfix/releases/latest/download/manifest.prod.xml}}"
NAME="plsfix-manifest.xml"
EXCEL="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
PPT="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

TMP="$(mktemp "${TMPDIR:-/tmp}/plsfix-manifest.XXXXXX")"
trap 'rm -f "$TMP"' EXIT
echo "Downloading the pls,fix manifest from $URL"
curl -fsSL "$URL" -o "$TMP"
if ! grep -q '<OfficeApp' "$TMP"; then
  echo "That file is not an Office add-in manifest: $URL" >&2
  exit 1
fi
VERSION="$(sed -n 's/.*<Version>\([0-9.]*\)<\/Version>.*/\1/p' "$TMP" | head -1)"

for DIR in "$EXCEL" "$PPT"; do
  mkdir -p "$DIR"
  cp "$TMP" "$DIR/$NAME"
done

echo "pls,fix ${VERSION:-} is installed for Excel and PowerPoint."
echo "Quit and reopen Excel and PowerPoint: the pls,fix tab appears on the ribbon,"
echo "Ctrl+Shift+M opens the pane. To remove it: plsfix-uninstall-mac.command."
