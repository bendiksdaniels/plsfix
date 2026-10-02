// A PNG that is only a signature plus an IHDR chunk: enough for pngSize(),
// stable text for equality checks. Not a decodable image. `padding` appends
// that many bytes past the header, which is how a test builds a picture big
// enough to meet the relay's response cap.
export function fakePng(width: number, height: number, padding = 0): string {
  const bytes = new Uint8Array(8 + 4 + 4 + 13 + 4 + padding);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  const view = new DataView(bytes.buffer);
  view.setUint32(8, 13);
  bytes.set([0x49, 0x48, 0x44, 0x52], 12); // "IHDR"
  view.setUint32(16, width);
  view.setUint32(20, height);
  bytes.set([8, 6, 0, 0, 0], 24); // bit depth, colour type, compression, filter, interlace
  view.setUint32(29, 0); // CRC not validated by pngSize
  return base64(bytes);
}

// btoa wants a binary string, and String.fromCharCode(...bytes) is a spread: a
// padded picture is megabytes long, so it goes in chunks the call stack holds.
const CHUNK = 0x8000;

function base64(bytes: Uint8Array): string {
  let binary = "";
  for (let at = 0; at < bytes.length; at += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(at, at + CHUNK));
  }
  return btoa(binary);
}
