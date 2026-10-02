// Attacks decodeBundle/readPastedBundle (src/link/bundle.ts) with hostile
// input: the invariant is "foreign text never throws here; it answers a
// reason instead" - targets the guard lines bundle.test.ts's happy path
// never reaches: a bad `v`, a non-object item, a bad date, a missing blob.

import { describe, expect, it } from "vitest";
import {
  decodeBundle,
  encodeBundle,
  readPastedBundle,
  type Bundle,
} from "../../src/link/bundle";

const ID = "0123456789abcdef0123456789abcdef";
const ID2 = "fedcba9876543210fedcba9876543210";

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

// Both take already-serialised JSON array text (never a JS value: passing
// one through JSON.stringify here would re-quote it as a single string
// field instead of splicing it in as the array literal every test below
// means to build) and splice it straight into a minimal envelope, so a
// refusal is pinned to the one guard that must catch it, never to
// listOf's earlier !Array.isArray check refusing a stringified decoy.
function withLinks(linksJson: string): string {
  return `{"plsfix":"links","v":1,"links":${linksJson},"inbox":[]}`;
}
function withInbox(inboxJson: string): string {
  return `{"plsfix":"links","v":1,"links":[],"inbox":${inboxJson}}`;
}

describe("decodeBundle: the envelope", () => {
  it.each([
    [
      "v missing",
      '{"plsfix":"links","links":[],"inbox":[{"id":"' +
        ID +
        '","createdAt":1,"blob":"AQ"}]}',
    ],
    [
      "v a string",
      '{"plsfix":"links","v":"1","links":[],"inbox":[{"id":"' +
        ID +
        '","createdAt":1,"blob":"AQ"}]}',
    ],
    [
      "v null",
      '{"plsfix":"links","v":null,"links":[],"inbox":[{"id":"' +
        ID +
        '","createdAt":1,"blob":"AQ"}]}',
    ],
    [
      "v zero",
      '{"plsfix":"links","v":0,"links":[],"inbox":[{"id":"' +
        ID +
        '","createdAt":1,"blob":"AQ"}]}',
    ],
    [
      "v negative",
      '{"plsfix":"links","v":-1,"links":[],"inbox":[{"id":"' +
        ID +
        '","createdAt":1,"blob":"AQ"}]}',
    ],
  ])("answers notBundle when %s", (_name, text) => {
    expect(decodeBundle(text)).toEqual({ ok: false, reason: "notBundle" });
  });

  it.each([1.5, 2, 9999])(
    "answers newerVersion for any v above the one this add-in knows (%s)",
    (v) => {
      const text = encodeBundle(sample).replace('"v":1', `"v":${String(v)}`);
      expect(decodeBundle(text)).toEqual({ ok: false, reason: "newerVersion" });
    },
  );

  it("answers notBundle for a top-level value that is not an object at all", () => {
    for (const text of ["null", "42", "true", '"a string"', "[]"]) {
      expect(decodeBundle(text)).toEqual({ ok: false, reason: "notBundle" });
    }
  });

  it("answers notBundle when links or inbox is present but not an array", () => {
    const linksNotArray = '{"plsfix":"links","v":1,"links":{},"inbox":[]}';
    const inboxNotArray = '{"plsfix":"links","v":1,"links":[],"inbox":"nope"}';
    expect(decodeBundle(linksNotArray)).toEqual({
      ok: false,
      reason: "notBundle",
    });
    expect(decodeBundle(inboxNotArray)).toEqual({
      ok: false,
      reason: "notBundle",
    });
  });
});

describe("decodeBundle: a links[] item that is not a well-formed BundleLink", () => {
  it.each([
    ["not an object", "42"],
    ["null", "null"],
    ["a string", '"nope"'],
    ["an array", "[]"],
  ])("refuses when the item is %s", (_name, item) => {
    expect(decodeBundle(withLinks(`[${item}]`)).ok).toBe(false);
  });

  it("refuses a link with no sentAt", () => {
    const item = `{"id":"${ID}","rev":1,"blob":"AQID"}`;
    expect(decodeBundle(withLinks(`[${item}]`)).ok).toBe(false);
  });

  it("refuses a link whose sentAt is not finite", () => {
    for (const sentAt of ['"yesterday"', "null", "NaN", "Infinity"]) {
      const text = withLinks(
        `[{"id":"${ID}","rev":1,"sentAt":${sentAt === "NaN" || sentAt === "Infinity" ? "1e999" : sentAt},"blob":"AQID"}]`,
      );
      expect(decodeBundle(text).ok).toBe(false);
    }
  });

  it("refuses a link whose blob is missing or empty", () => {
    const missing = `{"id":"${ID}","rev":1,"sentAt":1}`;
    const empty = `{"id":"${ID}","rev":1,"sentAt":1,"blob":""}`;
    const notString = `{"id":"${ID}","rev":1,"sentAt":1,"blob":7}`;
    expect(decodeBundle(withLinks(`[${missing}]`)).ok).toBe(false);
    expect(decodeBundle(withLinks(`[${empty}]`)).ok).toBe(false);
    expect(decodeBundle(withLinks(`[${notString}]`)).ok).toBe(false);
  });

  it("refuses one bad link even when every other field of the bundle is otherwise fine", () => {
    // listOf refuses the whole array on the first bad item: a bundle is
    // atomic, never half-applied.
    const good = `{"id":"${ID}","rev":1,"sentAt":1,"blob":"AQID"}`;
    const bad = `{"id":"${ID2}","rev":0,"sentAt":1,"blob":"AQID"}`; // rev < 1
    expect(decodeBundle(withLinks(`[${good},${bad}]`)).ok).toBe(false);
  });
});

describe("decodeBundle: an inbox[] item that is not a well-formed BundleInboxRow", () => {
  it.each([
    ["not an object", "42"],
    ["null", "null"],
  ])("refuses when the item is %s", (_name, item) => {
    expect(decodeBundle(withInbox(`[${item}]`)).ok).toBe(false);
  });

  it("refuses a row with a bad id", () => {
    const text = withInbox(`[{"id":"not-32-hex","createdAt":1,"blob":"AQ"}]`);
    expect(decodeBundle(text).ok).toBe(false);
  });

  it("refuses a row whose createdAt is not finite", () => {
    const text = withInbox(`[{"id":"${ID}","createdAt":"nope","blob":"AQ"}]`);
    expect(decodeBundle(text).ok).toBe(false);
  });

  it("refuses a row whose blob is missing or empty", () => {
    const missing = withInbox(`[{"id":"${ID}","createdAt":1}]`);
    const empty = withInbox(`[{"id":"${ID}","createdAt":1,"blob":""}]`);
    expect(decodeBundle(missing).ok).toBe(false);
    expect(decodeBundle(empty).ok).toBe(false);
  });
});

describe("readPastedBundle: the HTML carrier", () => {
  it("answers notBundle when the attribute value has a character the carrier regex refuses", () => {
    // "@" falls outside [A-Za-z0-9_-], so CARRIER never matches this
    // attribute at all: the read falls through to the plain-text branch
    // (readPastedBundle's OTHER path), not the try/catch around the base64
    // decode - covered separately below.
    const html = '<span data-plsfix-links="not@@valid">pls,fix links</span>';
    expect(readPastedBundle(html, "")).toEqual({
      ok: false,
      reason: "notBundle",
    });
  });

  it("answers notBundle when the carrier value IS in the base64url alphabet but decodes to nothing valid", () => {
    // A single character passes CARRIER's [A-Za-z0-9_-]+ and BASE64URL's own
    // regex, but pads to "A===", which atob() throws on ("Invalid
    // character"): the one path to bundle.ts's own catch around
    // fromBase64Url(carried), as opposed to the regex-mismatch case above.
    const html = '<span data-plsfix-links="A">pls,fix links</span>';
    expect(readPastedBundle(html, "")).toEqual({
      ok: false,
      reason: "notBundle",
    });
  });

  it("answers notBundle when the carrier decodes to base64url junk, not JSON", () => {
    // "aGVsbG8" is valid base64url for the plain word "hello".
    const html = '<span data-plsfix-links="aGVsbG8">pls,fix links</span>';
    expect(readPastedBundle(html, "")).toEqual({
      ok: false,
      reason: "notBundle",
    });
  });

  it("falls through to the plain-text bundle when there is no carrier attribute at all", () => {
    const json = encodeBundle(sample);
    expect(readPastedBundle("<b>no carrier here</b>", json)).toEqual({
      ok: true,
      bundle: sample,
    });
  });

  it("never throws on any of a wide sweep of garbage html/plain pairs", () => {
    const garbage = [
      "",
      "null",
      "undefined",
      "{}",
      "[]",
      "\u0000\u0001",
      "🎉".repeat(500),
      "<span data-plsfix-links>no value</span>",
      '<span data-plsfix-links="">empty</span>',
      "a".repeat(50_000),
    ];
    for (const html of garbage) {
      for (const plain of garbage) {
        expect(() => readPastedBundle(html, plain)).not.toThrow();
      }
    }
  });
});

describe("repeats: decoding the same malformed text three times", () => {
  it("answers exactly the same refusal every time - no hidden state accumulates", () => {
    const text = withLinks(
      '[{"id":"' + ID + '","rev":0,"sentAt":1,"blob":"AQ"}]',
    );
    const first = decodeBundle(text);
    const second = decodeBundle(text);
    const third = decodeBundle(text);
    expect(first).toEqual({ ok: false, reason: "notBundle" });
    expect(second).toEqual(first);
    expect(third).toEqual(first);
  });
});
