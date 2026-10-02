#!/bin/sh
# Removes pls,fix from Excel and PowerPoint on this Mac: deletes the sideloaded
# manifest the installer placed. Office forgets the add-in on its next launch;
# nothing else was ever installed.
# Usage: curl -fsSL <this file's URL> | sh        (or double-click it)
set -eu

NAME="plsfix-manifest.xml"
EXCEL="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
PPT="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

REMOVED=0
for DIR in "$EXCEL" "$PPT"; do
  if [ -f "$DIR/$NAME" ]; then
    rm -f "$DIR/$NAME"
    REMOVED=$((REMOVED + 1))
  fi
done

if [ "$REMOVED" -eq 0 ]; then
  echo "pls,fix was not installed by the installer on this Mac (nothing to remove)."
else
  echo "pls,fix is removed from $REMOVED app folder(s). Quit and reopen Excel and PowerPoint."
  echo "If the tab lingers, clear the Office cache: ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/ (Excel)"
  echo "and the com.microsoft.Powerpoint twin, then relaunch."
fi
