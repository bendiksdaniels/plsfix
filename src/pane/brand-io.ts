// The palette as a file: the JSON the Copy button hands over, and the rule
// that decides whether a picked file is a palette at all. Owns no state and
// touches no control - the Brand tab wires both to its own two inputs. The
// invariant: a file that says nothing about a palette is never applied as one.

import { serializeSettings, type BrandSettings } from "../settings";
import { DEFAULT_SETTINGS, getActiveSettings, parsePalette } from "../settings";
import { toast } from "./shared";

// A JSON file carrying none of the palette's own keys still parses into the
// shipped defaults, so importing somebody else's file - a package.json, an
// export from another tool - would wipe the modeller's colours under a
// "Palette imported" toast. The key list is the settings shape itself, so the
// two can never drift apart.
function isPaletteFile(json: string): boolean {
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch {
    return false;
  }
  if (typeof raw !== "object" || raw === null) return false;
  return Object.keys(DEFAULT_SETTINGS).some((key) => key in raw);
}

/** The palette a picked file carries, or null when it carries none. */
export function readPaletteFile(json: string): BrandSettings | null {
  return isPaletteFile(json) ? parsePalette(json) : null;
}

/**
 * The active palette onto the clipboard. Office webviews can deny the
 * Clipboard API, so a hidden textarea is the fallback; either way the
 * modeller gets one sentence and this never throws.
 */
export async function copyPaletteJson(): Promise<void> {
  const json = serializeSettings(getActiveSettings());
  try {
    await navigator.clipboard.writeText(json);
    toast.show("Palette JSON copied");
    return;
  } catch {
    // Fall through to the textarea.
  }

  const area = document.createElement("textarea");
  area.value = json;
  document.body.append(area);
  area.select();
  // execCommand is deprecated and a webview may have dropped it; the last
  // resort still owes a sentence, not a rejection nobody awaits.
  let copied = false;
  try {
    copied = document.execCommand("copy");
  } catch {
    copied = false;
  }
  area.remove();
  toast.show(
    copied ? "Palette JSON copied" : "Copy failed",
    copied ? "success" : "error",
  );
}
