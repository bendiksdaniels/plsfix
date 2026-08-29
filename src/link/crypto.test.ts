import { describe, expect, it } from "vitest";
import {
  deriveLinkKeys,
  fromBase64Url,
  hkdf,
  newToken,
  open,
  seal,
  sha256Hex,
  toBase64Url,
} from "./crypto";

const bytes = (...values: number[]) => new Uint8Array(values);
const hex = (data: Uint8Array) =>
  [...data].map((byte) => byte.toString(16).padStart(2, "0")).join("");

describe("base64url", () => {
  it("round-trips without padding", () => {
    const text = toBase64Url(bytes(251, 255, 191));
    expect(text).toBe("-_-_");
    expect(fromBase64Url(text)).toEqual(bytes(251, 255, 191));
    expect(() => fromBase64Url("not*valid")).toThrow();
  });
  it("survives a payload wider than the encoder's chunk", () => {
    const png = new Uint8Array(new ArrayBuffer(70000));
    for (let index = 0; index < png.length; index += 1) {
      png[index] = (index * 31) % 256;
    }
    expect(fromBase64Url(toBase64Url(png))).toEqual(png);
  });
});

describe("hkdf", () => {
  it("is deterministic and info-separated", async () => {
    const secret = bytes(...Array(32).fill(7));
    const a = await hkdf(secret, "plsfix-link-enc");
    const b = await hkdf(secret, "plsfix-link-enc");
    const c = await hkdf(secret, "plsfix-link-auth");
    expect(a).toEqual(b);
    expect(a).not.toEqual(c);
    expect(a).toHaveLength(32);
  });
  it("matches RFC 5869 test case 1 shape (32 bytes from 22-byte IKM)", async () => {
    const ikm = bytes(...Array(22).fill(0x0b));
    expect(await hkdf(ikm, "", 42)).toHaveLength(42);
  });
  it("matches the RFC 5869 test case 1 output for its salt and info", async () => {
    const ikm = bytes(...Array(22).fill(0x0b));
    const salt = new Uint8Array([
      0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b,
      0x0c,
    ]);
    const info = new Uint8Array([
      0xf0, 0xf1, 0xf2, 0xf3, 0xf4, 0xf5, 0xf6, 0xf7, 0xf8, 0xf9,
    ]);
    expect(hex(await hkdf(ikm, info, 42, salt))).toBe(
      "3cb25f25faacd57a90434f64d0362f2a" +
        "2d2d0a90cf1a5a4c5db02d56ecc4c5bf34007208d5b887185865",
    );
  });
});

describe("seal/open", () => {
  it("round-trips and binds the AAD", async () => {
    const key = await hkdf(bytes(...Array(32).fill(1)), "plsfix-link-enc");
    const plain = new TextEncoder().encode("hello");
    const blob = await seal(key, "link-1", plain);
    expect(blob.length).toBe(12 + plain.length + 16);
    expect(await open(key, "link-1", blob)).toEqual(plain);
    await expect(open(key, "link-2", blob)).rejects.toThrow(/decrypt/);
    const other = await hkdf(bytes(...Array(32).fill(2)), "plsfix-link-enc");
    await expect(open(other, "link-1", blob)).rejects.toThrow(/decrypt/);
  });
  it("uses a fresh IV per call", async () => {
    const key = await hkdf(bytes(...Array(32).fill(1)), "plsfix-link-enc");
    const a = await seal(key, "x", bytes(1));
    const b = await seal(key, "x", bytes(1));
    expect(a.slice(0, 12)).not.toEqual(b.slice(0, 12));
  });
});

describe("tokens", () => {
  it("are 43 base64url chars and derive distinct enc/auth", async () => {
    const token = newToken();
    expect(token).toMatch(/^[A-Za-z0-9_-]{43}$/);
    const keys = await deriveLinkKeys(token);
    expect(keys.enc).toHaveLength(32);
    expect(keys.auth).toMatch(/^[A-Za-z0-9_-]{43}$/);
    expect(toBase64Url(keys.enc)).not.toBe(keys.auth);
  });
});

describe("sha256Hex", () => {
  it("hashes strings and bytes", async () => {
    const digest =
      "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad";
    expect(await sha256Hex("abc")).toBe(digest);
    expect(await sha256Hex(new TextEncoder().encode("abc"))).toBe(digest);
  });
});

describe("key length guards", () => {
  it("rejects tokens that are not 32 bytes and keys that are not 32 bytes", async () => {
    await expect(deriveLinkKeys("")).rejects.toThrow(/32 bytes/);
    await expect(deriveLinkKeys(toBase64Url(bytes(1, 2, 3)))).rejects.toThrow(
      /32 bytes/,
    );
    const short = bytes(...Array(16).fill(9));
    await expect(seal(short, "x", bytes(1))).rejects.toThrow(/32 bytes/);
  });
});
