// Clipboard copy shared by the toast's "Copy details" button, the Excel
// Links tab's key copy and the model check report. Prefers the async
// Clipboard API; Office webviews can deny it, so a hidden textarea +
// execCommand("copy") is the fallback - either way this never throws, but it
// answers whether the text actually reached the clipboard so a caller can
// tell the modeller when it did not.

// The one sentence every copy button answers with when neither path landed
// the text on the clipboard, shared rather than restated per caller.
export const COPY_FAILED_MESSAGE =
  "Copy failed: select the text and copy it by hand.";

export async function copyText(text: string): Promise<boolean> {
  if (navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      // Fall through to the textarea.
    }
  }
  return copyViaHiddenTextarea(text);
}

// execCommand("copy") only acts on a selected, focused element, so this
// needs a real (if invisible) textarea in the document; a host with neither
// API (jsdom, an old webview) just leaves the clipboard untouched and answers
// false, same as execCommand itself returning false.
function copyViaHiddenTextarea(text: string): boolean {
  const area = document.createElement("textarea");
  area.value = text;
  area.style.position = "fixed";
  area.style.opacity = "0";
  document.body.append(area);
  area.select();
  try {
    return document.execCommand("copy");
  } catch {
    // Not implemented in this host.
    return false;
  } finally {
    area.remove();
  }
}
