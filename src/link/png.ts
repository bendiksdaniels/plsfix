// PNG size reader: the fixed 8-byte signature plus the IHDR chunk's
// fixed-offset width and height fields, nothing else. No decoding, no image
// library - callers only need pixel dimensions before a picture from
// Range.getImage or the relay is ever drawn.

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const IHDR_WIDTH_OFFSET = 16;
const IHDR_HEIGHT_OFFSET = 20;
const MIN_PNG_BYTES = IHDR_HEIGHT_OFFSET + 4; // sig + length + "IHDR" + w + h

function hasPngSignature(bytes: Uint8Array): boolean {
  return (
    bytes.length >= MIN_PNG_BYTES &&
    PNG_SIGNATURE.every((byte, index) => bytes[index] === byte)
  );
}

export function pngSize(bytes: Uint8Array): { width: number; height: number } {
  if (!hasPngSignature(bytes)) throw new Error("pngSize: not a PNG");

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return {
    width: view.getUint32(IHDR_WIDTH_OFFSET),
    height: view.getUint32(IHDR_HEIGHT_OFFSET),
  };
}

export function base64ToBytes(base64: string): Uint8Array {
  const commaIndex = base64.indexOf(",");
  const raw =
    base64.startsWith("data:") && commaIndex >= 0
      ? base64.slice(commaIndex + 1)
      : base64;

  const binary = atob(raw);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}
