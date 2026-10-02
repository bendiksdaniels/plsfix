// The link core's security and codec edges: that a token's two key branches
// and a workspace's three are separate secrets, that no seal ever reuses a
// nonce, that base64url round trips whatever length and leading zeros a key or
// a picture gives it, and that every decoder refuses garbage with a named
// error rather than a typed value.
// Invariant: one secret, one purpose - a key derived for one branch never
// opens another's blob.
import { describe, expect, it } from "vitest";
import {
  decodeBase64,
  deriveLinkKeys,
  fromBase64Url,
  hkdf,
  newToken,
  open,
  randomBytes,
  seal,
  sha256Hex,
  toBase64Url,
} from "./crypto";
import { deriveWorkspace } from "./workspace";
import {
  decodeInboxItem,
  decodePayload,
  decodeTag,
  encodeInboxItem,
  encodePayload,
  encodeTag,
  tryDecodeRegistry,
  TAG_VALUE_MAX,
  type Payload,
} from "./model";
import { Debouncer, type Clock, type Timer } from "./debounce";
import { pngSize } from "./png";

const SRC = { workbook: "M.xlsx", sheet: "P&L", ref: "A1:B2", anchor: "A1" };

function payload(): Payload {
  return {
    v: 1,
    kind: "text",
    text: "42",
    src: SRC,
    pushedAt: "2026-09-08T10:00:00.000Z",
    hash: "h",
  };
}

describe("key derivation keeps its branches apart", () => {
  // The same 32 bytes are a link token in Excel's registry and a pairing key
  // in the host store. Both derive an enc key and a bearer, and nothing may
  // line up: one leaked bearer would otherwise be another branch's enc key.
  it("derives five different keys from one secret", async () => {
    const secret = randomBytes(32);
    const link = await deriveLinkKeys(toBase64Url(secret));
    const ws = await deriveWorkspace(secret);
    const derived = [
      toBase64Url(link.enc),
      link.auth,
      ws.id,
      ws.auth,
      toBase64Url(ws.enc),
    ];
    expect(new Set(derived).size).toBe(derived.length);
    // And none of them is the secret itself, which never leaves the host.
    expect(derived).not.toContain(ws.exportKey);
  });

  it("cannot open a link blob with the workspace key, or the other way", async () => {
    const secret = randomBytes(32);
    const link = await deriveLinkKeys(toBase64Url(secret));
    const ws = await deriveWorkspace(secret);
    const sealed = await seal(link.enc, "linkid", encodePayload(payload()));
    await expect(open(ws.enc, "linkid", sealed)).rejects.toThrow(
      "cannot decrypt",
    );
    const note = await seal(
      ws.enc,
      ws.id,
      encodeInboxItem({
        id: "a".repeat(32),
        token: ws.exportKey,
        kind: "text",
        label: "l",
        src: SRC,
        createdAt: "now",
      }),
    );
    await expect(open(link.enc, ws.id, note)).rejects.toThrow("cannot decrypt");
    // The right key with the wrong link id is refused too: the AAD binds the
    // blob to the row it was stored under.
    await expect(open(link.enc, "another", sealed)).rejects.toThrow(
      "cannot decrypt",
    );
  });

  it("gives every seal its own 96-bit nonce", async () => {
    const key = randomBytes(32);
    const plain = encodePayload(payload());
    const blobs = [];
    for (let round = 0; round < 64; round += 1) {
      blobs.push(await seal(key, "id", plain));
    }
    const nonces = new Set(
      blobs.map((blob) => toBase64Url(blob.subarray(0, 12))),
    );
    expect(nonces.size).toBe(blobs.length);
    // The same picture re-pushed is a different blob every time, which is what
    // keeps a repeated push from being recognisable on the wire.
    expect(new Set(blobs.map(toBase64Url)).size).toBe(blobs.length);
  });

  it("refuses a key or a length WebCrypto would take silently", async () => {
    await expect(hkdf(randomBytes(32), "x", 0)).rejects.toThrow("bad length");
    await expect(hkdf(randomBytes(32), "x", 1.5)).rejects.toThrow("bad length");
    await expect(hkdf(randomBytes(32), "x", 255 * 32 + 1)).rejects.toThrow(
      "bad length",
    );
    await expect(deriveLinkKeys(toBase64Url(randomBytes(31)))).rejects.toThrow(
      "32 bytes",
    );
    // A blob too short to hold an IV and a tag cannot be authentic.
    await expect(open(randomBytes(32), "id", randomBytes(27))).rejects.toThrow(
      "cannot decrypt",
    );
  });
});

describe("base64url", () => {
  // A token, a bearer and a sealed picture all travel through this: every
  // length mod 3 has to survive, and so do the leading zero bytes one token in
  // 256 starts with.
  it("round trips every length, leading zeros included", () => {
    for (let length = 0; length <= 36; length += 1) {
      const bytes = new Uint8Array(length);
      for (let index = 2; index < length; index += 1)
        bytes[index] = (index * 37) % 256;
      const text = toBase64Url(bytes);
      expect(text, `length ${String(length)}`).toMatch(/^[A-Za-z0-9_-]*$/);
      expect(fromBase64Url(text), `length ${String(length)}`).toEqual(bytes);
    }
  });

  it("round trips a picture-sized array across the chunk boundary", () => {
    const bytes = new Uint8Array(0x8000 * 2 + 5);
    for (let index = 0; index < bytes.length; index += 1)
      bytes[index] = index % 256;
    expect(fromBase64Url(toBase64Url(bytes))).toEqual(bytes);
  });

  it("refuses text that is not base64url without quoting it back", () => {
    for (const text of ["A", "AB==", "a*b", " AA ", "AA\n", "AAAA/", "AA+A"]) {
      expect(() => fromBase64Url(text), text).toThrow(
        "fromBase64Url: not base64url",
      );
      expect(() => fromBase64Url(text), text).not.toThrow(text);
    }
    expect(() => decodeBase64("====")).toThrow();
  });

  it("makes a 43-character token and a hex digest of a fixed width", async () => {
    expect(newToken()).toHaveLength(43);
    expect(await sha256Hex("x")).toHaveLength(64);
    expect(await sha256Hex(new Uint8Array([1]))).toMatch(/^[0-9a-f]{64}$/);
  });
});

describe("codecs refuse what they cannot type", () => {
  it("names the stage when bytes are not JSON at all", () => {
    const junk = new Uint8Array([0xff, 0xfe, 0x00]);
    expect(() => decodePayload(junk)).toThrow("decodePayload: not valid JSON");
    expect(() => decodeInboxItem(junk)).toThrow(
      "decodeInboxItem: not valid JSON",
    );
  });

  it("refuses JSON of the wrong shape", () => {
    const bytes = (value: unknown) =>
      new TextEncoder().encode(JSON.stringify(value));
    expect(() => decodePayload(bytes({ v: 1, kind: "picture" }))).toThrow(
      "not a link payload",
    );
    expect(() => decodeInboxItem(bytes({ id: "a", token: "t" }))).toThrow(
      "not an inbox item",
    );
    // A table whose grid does not match its own counts is half a payload.
    const grid = {
      ...payload(),
      kind: "table",
      rows: 2,
      cols: 1,
      cells: [[{ t: "a" }]],
      widths: [10],
    };
    expect(() => decodePayload(bytes(grid))).toThrow("not a link payload");
  });

  it("tells an unreadable registry from an absent one", () => {
    expect(tryDecodeRegistry(null)).toBeNull();
    expect(tryDecodeRegistry("{")).toBeNull();
    expect(tryDecodeRegistry(JSON.stringify({ v: 2, links: [] }))).toBeNull();
    expect(tryDecodeRegistry(JSON.stringify({ v: 1, links: [] }))).toEqual({
      v: 1,
      links: [],
    });
  });

  it("refuses a tag longer than a shape can hold, and reads one back", () => {
    const tag = {
      v: 1 as const,
      id: "a".repeat(32),
      kind: "range" as const,
      rev: 3,
      src: SRC,
      pushedAt: "2026-09-08T10:00:00.000Z",
    };
    expect(decodeTag(encodeTag(tag))).toEqual(tag);
    const long = {
      ...tag,
      src: { ...SRC, workbook: "W".repeat(TAG_VALUE_MAX) },
    };
    expect(() => encodeTag(long)).toThrow(String(TAG_VALUE_MAX));
    expect(decodeTag(null)).toBeNull();
    expect(decodeTag("{}")).toBeNull();
  });

  it("refuses bytes that are not a PNG rather than reading a size out of them", () => {
    expect(() => pngSize(new Uint8Array(4))).toThrow("pngSize: not a PNG");
    expect(() => pngSize(new Uint8Array(64))).toThrow("pngSize: not a PNG");
  });
});

describe("Debouncer", () => {
  // A clock that fires the timer before anything is due - a host that rounds
  // its timeouts down - must re-arm and wait out the remainder, never flush an
  // empty push.
  it("waits out the remainder when its timer fires early", () => {
    const flushed: string[][] = [];
    let at = 0;
    // A stack, because the debouncer always cancels before it arms again: the
    // last entry is the timer that is actually live.
    const armed: (() => void)[] = [];
    const clock: Clock = {
      now: () => at,
      after: (_ms, run): Timer => {
        armed.push(run);
        return {
          cancel: () => {
            armed.pop();
          },
        };
      },
    };
    const fire = (): void => {
      armed[armed.length - 1]?.();
    };
    const debouncer = new Debouncer(1000, (keys) => flushed.push(keys), clock);
    debouncer.touch(["Sheet1"]);
    at = 999;
    fire();
    expect(flushed).toEqual([]);
    expect(debouncer.pending()).toEqual(["Sheet1"]);
    at = 1000;
    fire();
    expect(flushed).toEqual([["Sheet1"]]);
    expect(debouncer.pending()).toEqual([]);
  });
});
