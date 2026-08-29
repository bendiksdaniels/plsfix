// Link crypto: the add-in's only WebCrypto surface - base64url, HKDF-SHA256,
// AES-256-GCM seal/open (12-byte IV prefix, AAD bound to the link or workspace
// id) and the 32-byte link token. Pure: no Office.js, no network, no storage.
// Invariant: no error message ever carries key material, only the stage.

const webcrypto = globalThis.crypto;
const ENCODER = new TextEncoder();
const BASE64URL = /^[A-Za-z0-9_-]*$/;
const EMPTY_SALT = new Uint8Array(0);
const KEY_BYTES = 32;
const IV_BYTES = 12;
const TAG_BYTES = 16;
// Largest argument run String.fromCharCode takes without risking a stack blow-up.
const CHUNK = 0x8000;

// WebCrypto takes a BufferSource, but a caller's Uint8Array is typed over
// ArrayBufferLike. Re-view the same bytes over their ArrayBuffer; only the
// SharedArrayBuffer case, which no Office host produces, costs a copy.
function source(data: Uint8Array): Uint8Array<ArrayBuffer> {
  const buffer = data.buffer;
  if (buffer instanceof ArrayBuffer) {
    return new Uint8Array(buffer, data.byteOffset, data.byteLength);
  }
  const copy = new Uint8Array(new ArrayBuffer(data.byteLength));
  copy.set(data);
  return copy;
}

export function randomBytes(length: number): Uint8Array {
  const out = new Uint8Array(new ArrayBuffer(length));
  webcrypto.getRandomValues(out);
  return out;
}

// Chunked, because a link payload is a whole PNG: one call per byte would turn
// a few megabytes into millions of string joins.
export function toBase64Url(bytes: Uint8Array): string {
  let binary = "";
  for (let index = 0; index < bytes.length; index += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(index, index + CHUNK));
  }
  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

// The one atob-to-bytes loop in the link core: the standard alphabet, no
// validation and no prefix handling, so both decoders (this file's base64url
// and png.ts's data: URL) own only what actually differs between them.
// Throws whatever atob throws on text outside the alphabet.
export function decodeBase64(base64: string): Uint8Array {
  const binary = atob(base64);
  const out = new Uint8Array(new ArrayBuffer(binary.length));
  for (let index = 0; index < binary.length; index += 1) {
    out[index] = binary.charCodeAt(index);
  }
  return out;
}

// Rejects anything outside the alphabet before atob sees it; the text itself is
// never quoted back, because a token or a workspace key is exactly this text.
export function fromBase64Url(text: string): Uint8Array {
  if (!BASE64URL.test(text)) throw new Error("fromBase64Url: not base64url");
  const base64 =
    text.replace(/-/g, "+").replace(/_/g, "/") +
    "=".repeat((4 - (text.length % 4)) % 4);
  try {
    return decodeBase64(base64);
  } catch {
    throw new Error("fromBase64Url: not base64url");
  }
}

// Info separates the purposes of one secret ("plsfix-link-enc" vs "plsfix-link-auth").
// Bytes are accepted so published test vectors, whose info is not UTF-8, are
// checkable; the salt defaults to empty, which HMAC pads to the RFC 5869 zeros.
export async function hkdf(
  secret: Uint8Array,
  info: string | Uint8Array,
  length = KEY_BYTES,
  salt: Uint8Array = EMPTY_SALT,
): Promise<Uint8Array> {
  if (!Number.isInteger(length) || length < 1 || length > 255 * 32) {
    throw new Error(`hkdf: bad length ${length}`);
  }
  const key = await webcrypto.subtle.importKey(
    "raw",
    source(secret),
    "HKDF",
    false,
    ["deriveBits"],
  );
  const bits = await webcrypto.subtle.deriveBits(
    {
      name: "HKDF",
      hash: "SHA-256",
      salt: source(salt),
      info: typeof info === "string" ? ENCODER.encode(info) : source(info),
    },
    key,
    length * 8,
  );
  return new Uint8Array(bits);
}

// blob = 12-byte random IV || AES-256-GCM ciphertext with its 16-byte tag.
export async function seal(
  key: Uint8Array,
  aad: string,
  plaintext: Uint8Array,
): Promise<Uint8Array> {
  const iv = source(randomBytes(IV_BYTES));
  const cipher = await webcrypto.subtle.encrypt(
    { name: "AES-GCM", iv, additionalData: ENCODER.encode(aad) },
    await aesKey(key, "encrypt"),
    source(plaintext),
  );
  const blob = new Uint8Array(new ArrayBuffer(IV_BYTES + cipher.byteLength));
  blob.set(iv, 0);
  blob.set(new Uint8Array(cipher), IV_BYTES);
  return blob;
}

// Wrong key, wrong AAD and a truncated blob are the same answer to a caller:
// there is nothing to show. The reason stays inside WebCrypto.
export async function open(
  key: Uint8Array,
  aad: string,
  blob: Uint8Array,
): Promise<Uint8Array> {
  if (blob.length < IV_BYTES + TAG_BYTES) {
    throw new Error("open: cannot decrypt");
  }
  try {
    const plain = await webcrypto.subtle.decrypt(
      {
        name: "AES-GCM",
        iv: source(blob.subarray(0, IV_BYTES)),
        additionalData: ENCODER.encode(aad),
      },
      await aesKey(key, "decrypt"),
      source(blob.subarray(IV_BYTES)),
    );
    return new Uint8Array(plain);
  } catch {
    throw new Error("open: cannot decrypt");
  }
}

export async function sha256Hex(data: Uint8Array | string): Promise<string> {
  const input = typeof data === "string" ? ENCODER.encode(data) : source(data);
  const digest = await webcrypto.subtle.digest("SHA-256", input);
  return [...new Uint8Array(digest)]
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

export interface LinkKeys {
  enc: Uint8Array;
  auth: string;
}

// The token is the whole secret of one link: it lives in the deck's PLSFIX_KEY tag
// and the workbook registry, never on the relay.
export function newToken(): string {
  return toBase64Url(randomBytes(KEY_BYTES));
}

export async function deriveLinkKeys(token: string): Promise<LinkKeys> {
  const secret = fromBase64Url(token);
  if (secret.length !== KEY_BYTES) {
    throw new Error(`deriveLinkKeys: token must be ${KEY_BYTES} bytes`);
  }
  return {
    enc: await hkdf(secret, "plsfix-link-enc"),
    auth: toBase64Url(await hkdf(secret, "plsfix-link-auth")),
  };
}

function aesKey(
  key: Uint8Array,
  usage: "encrypt" | "decrypt",
): Promise<CryptoKey> {
  if (key.length !== KEY_BYTES) {
    throw new Error(`aesKey: key must be ${KEY_BYTES} bytes`);
  }
  return webcrypto.subtle.importKey("raw", source(key), "AES-GCM", false, [
    usage,
  ]);
}
