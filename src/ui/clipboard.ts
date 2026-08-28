// Clipboard copy shared by the toast's "Copy details" button and the Excel
// Links tab's key copy. Prefers the async Clipboard API; Office webviews can
// deny it, so a hidden textarea + execCommand("copy") is the fallback -
// either way this never throws.

export async function copyText(text: string): Promise<void> {
  if (navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return;
    } catch {
      // Fall through to the textarea.
    }
  }
  copyViaHiddenTextarea(text);
}

// execCommand("copy") only acts on a selected, focused element, so this
// needs a real (if invisible) textarea in the document; a host with neither
// API (jsdom, an old webview) just leaves the clipboard untouched.
function copyViaHiddenTextarea(text: string): void {
  const area = document.createElement("textarea");
  area.value = text;
  area.style.position = "fixed";
  area.style.opacity = "0";
  document.body.append(area);
  area.select();
  try {
    document.execCommand("copy");
  } catch {
    // Not implemented in this host.
  } finally {
    area.remove();
  }
}
