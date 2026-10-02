// src/link/bundle.test.ts
// The clipboard bundle: a round trip, the guards foreign text meets, and the
// HTML carrier a paste reads first.
import { describe, expect, it } from "vitest";
import {
  BUNDLE_SENTENCE,
  bundleHtml,
  decodeBundle,
  encodeBundle,
  readPastedBundle,
  type Bundle,
} from "./bundle";

const ID = "0123456789abcdef0123456789abcdef";
const sample: Bundle = {
  links: [
    {
      id: ID,
      rev: 2 ** 40 + 1,
      sentAt: 1_790_000_000_000,
      blob: new Uint8Array([1, 2, 3]),
    },
  ],
  inbox: [
    { id: ID, createdAt: 1_790_000_000_000, blob: new Uint8Array([9, 8]) },
  ],
};

describe("encodeBundle / decodeBundle", () => {
  it("round-trips every field", () => {
    const read = decodeBundle(encodeBundle(sample));
    expect(read).toEqual({ ok: true, bundle: sample });
  });

  it("answers notBundle for text that is not one", () => {
    for (const text of ["", "hello", "{}", "[1]", '{"plsfix":"other","v":1}']) {
      expect(decodeBundle(text)).toEqual({ ok: false, reason: "notBundle" });
    }
  });

  it("answers newerVersion for a bundle from a newer add-in", () => {
    const newer = encodeBundle(sample).replace('"v":1', '"v":2');
    expect(decodeBundle(newer)).toEqual({ ok: false, reason: "newerVersion" });
  });

  it("refuses a bad id, a bad rev or a blob that is not base64url", () => {
    const text = encodeBundle(sample);
    expect(decodeBundle(text.replace(ID, "not-an-id")).ok).toBe(false);
    expect(
      decodeBundle(text.replace(`"rev":${String(2 ** 40 + 1)}`, '"rev":-1')).ok,
    ).toBe(false);
    expect(decodeBundle(text.replace('"blob":"AQID"', '"blob":"@@"')).ok).toBe(
      false,
    );
  });

  it("refuses an empty bundle", () => {
    expect(decodeBundle(encodeBundle({ links: [], inbox: [] })).ok).toBe(false);
  });
});

describe("the clipboard carrier", () => {
  it("hides the bundle in one attribute and shows only the sentence", () => {
    const html = bundleHtml(encodeBundle(sample));
    expect(html.startsWith('<span data-plsfix-links="')).toBe(true);
    expect(html.endsWith(`">${BUNDLE_SENTENCE}</span>`)).toBe(true);
  });

  it("reads the HTML flavor first, the plain text second", () => {
    const json = encodeBundle(sample);
    expect(readPastedBundle(bundleHtml(json), BUNDLE_SENTENCE)).toEqual({
      ok: true,
      bundle: sample,
    });
    expect(readPastedBundle("", `  ${json}\n`)).toEqual({
      ok: true,
      bundle: sample,
    });
    expect(readPastedBundle("<b>hi</b>", "hi")).toEqual({
      ok: false,
      reason: "notBundle",
    });
  });

  it("survives a sanitizer that re-serialises the span around it", () => {
    const json = encodeBundle(sample);
    const wrapped = `<html><body><!--StartFragment-->${bundleHtml(json).replace("<span ", '<span style="color:black" ')}<!--EndFragment--></body></html>`;
    expect(readPastedBundle(wrapped, "").ok).toBe(true);
  });
});
