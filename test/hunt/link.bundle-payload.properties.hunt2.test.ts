// Pass-2 properties: src/link/bundle.ts and src/link/model.ts's codecs. A
// bundle, or a picture/table/text payload, decodes back to exactly what was
// encoded, and an over-cap table render is refused before sealing, on every
// render. The clipboard cap boundary is its own file (line count).

import fc from "fast-check";
import { afterEach, describe, expect, it, vi } from "vitest";
import { decodeBundle, encodeBundle } from "../../src/link/bundle";
import {
  anchorName,
  decodePayload,
  encodePayload,
  TABLE_MAX_COLS,
  TABLE_MAX_ROWS,
  TABLE_TOO_BIG,
  type PicturePayload,
  type Source,
  type TableCell,
  type TablePayload,
  type TextPayload,
} from "../../src/link/model";
import { createWorkspace, type KeyStore } from "../../src/link/workspace";
import { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";

enableStrictLoadSemantics();

const SEED = 20260927;
const RUNS = 200;
const TABLE_RUNS = 80;

function toHex(bytes: Uint8Array): string {
  return Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
}

const linkIdArb = fc.uint8Array({ minLength: 16, maxLength: 16 }).map(toHex);
const textArb = (maxLength: number) =>
  fc.string({ unit: "grapheme", maxLength });
const sourceArb: fc.Arbitrary<Source> = fc.record({
  workbook: textArb(30),
  sheet: textArb(30),
  ref: textArb(20),
  anchor: textArb(20),
});
const safeIntArb = fc.integer({ min: 0, max: 1_000_000 });

// -----------------------------------------------------------------------
// Bundle round trip
// -----------------------------------------------------------------------

// A real blob is always sealed ciphertext (12-byte IV + tag at minimum, per
// src/link/crypto.ts's seal()), so it is never actually empty; an empty
// blob is a documented, pass-1-pinned refusal
// (test/hunt/link.bundle-malformed.hunt.test.ts "refuses a link whose blob
// is missing or empty"), not a shape this round-trip property should generate.
const blobArb = fc.uint8Array({ minLength: 28, maxLength: 200 });
const bundleLinkArb = fc.record({
  id: linkIdArb,
  rev: fc.integer({ min: 1, max: Number.MAX_SAFE_INTEGER }),
  sentAt: fc.integer({ min: 0, max: Number.MAX_SAFE_INTEGER }),
  blob: blobArb,
});
const bundleRowArb = fc.record({
  id: linkIdArb,
  createdAt: fc.integer({ min: 0, max: Number.MAX_SAFE_INTEGER }),
  blob: blobArb,
});
const bundleArb = fc
  .record({
    links: fc.array(bundleLinkArb, { maxLength: 5 }),
    inbox: fc.array(bundleRowArb, { maxLength: 5 }),
  })
  .filter((bundle) => bundle.links.length + bundle.inbox.length > 0);

describe("bundle encode/decode round trip", () => {
  it("decodes back to exactly what was encoded, for any links and inbox rows", () => {
    fc.assert(
      fc.property(bundleArb, (bundle) => {
        const read = decodeBundle(encodeBundle(bundle));
        expect(read).toEqual({ ok: true, bundle });
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

// -----------------------------------------------------------------------
// Payload round trips
// -----------------------------------------------------------------------

const pictureArb: fc.Arbitrary<PicturePayload> = fc.record({
  v: fc.constant(1 as const),
  kind: fc.constant("picture" as const),
  mime: fc.constant("image/png" as const),
  width: safeIntArb,
  height: safeIntArb,
  png: textArb(50),
  src: sourceArb,
  pushedAt: textArb(24),
  hash: textArb(64),
});

describe("picture payload round trip", () => {
  it("decodes back to exactly what was encoded, for any random fields", () => {
    fc.assert(
      fc.property(pictureArb, (payload) => {
        expect(decodePayload(encodePayload(payload))).toEqual(payload);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

const tableCellArb: fc.Arbitrary<TableCell> = fc
  .record({
    t: textArb(20),
    b: fc.option(fc.constant(true as const), { nil: undefined }),
    i: fc.option(fc.constant(true as const), { nil: undefined }),
    c: fc.option(textArb(8), { nil: undefined }),
    f: fc.option(textArb(8), { nil: undefined }),
    a: fc.option(fc.constantFrom("l" as const, "c" as const, "r" as const), {
      nil: undefined,
    }),
    z: fc.option(fc.integer({ min: 1, max: 400 }), { nil: undefined }),
  })
  .map((cell) => {
    // Every optional key must be genuinely ABSENT when unset (never present
    // with value undefined), matching what the real encoder writes: model.ts
    // is documented to omit every default rather than spell it out.
    const out: TableCell = { t: cell.t };
    if (cell.b !== undefined) out.b = cell.b;
    if (cell.i !== undefined) out.i = cell.i;
    if (cell.c !== undefined) out.c = cell.c;
    if (cell.f !== undefined) out.f = cell.f;
    if (cell.a !== undefined) out.a = cell.a;
    if (cell.z !== undefined) out.z = cell.z;
    return out;
  });

// Grid dimensions and cells travel together, exactly the way link-table.ts
// hands renderTable's TableRender to payloadOf: rows/cols always agree with
// cells.length / cells[0].length.
const tableGridArb = fc
  .tuple(
    fc.integer({ min: 0, max: TABLE_MAX_ROWS }),
    fc.integer({ min: 0, max: TABLE_MAX_COLS }),
  )
  .chain(([rows, cols]) =>
    fc.record({
      rows: fc.constant(rows),
      cols: fc.constant(cols),
      cells: fc.array(
        fc.array(tableCellArb, { minLength: cols, maxLength: cols }),
        {
          minLength: rows,
          maxLength: rows,
        },
      ),
      widths: fc.array(fc.integer({ min: 1, max: 400 }), {
        minLength: cols,
        maxLength: cols,
      }),
    }),
  );

const tableArb: fc.Arbitrary<TablePayload> = fc
  .tuple(
    tableGridArb,
    sourceArb,
    textArb(24),
    textArb(64),
    fc.option(fc.constant(true as const), { nil: undefined }),
  )
  .map(([grid, src, pushedAt, hash, h]) => ({
    v: 1 as const,
    kind: "table" as const,
    ...grid,
    src,
    pushedAt,
    hash,
    ...(h === undefined ? {} : { h }),
  }));

describe("table payload round trip", () => {
  it("decodes back to exactly what was encoded, up to 60x20 with header flag and fills", () => {
    fc.assert(
      fc.property(tableArb, (payload) => {
        expect(decodePayload(encodePayload(payload))).toEqual(payload);
      }),
      { seed: SEED, numRuns: TABLE_RUNS },
    );
  });
});

const textPayloadArb: fc.Arbitrary<TextPayload> = fc.record({
  v: fc.constant(1 as const),
  kind: fc.constant("text" as const),
  text: textArb(500),
  src: sourceArb,
  pushedAt: textArb(24),
  hash: textArb(64),
});

describe("text payload round trip", () => {
  it("decodes back to exactly what was encoded, for any unicode text up to 500 chars", () => {
    fc.assert(
      fc.property(textPayloadArb, (payload) => {
        expect(decodePayload(encodePayload(payload))).toEqual(payload);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

// -----------------------------------------------------------------------
// Caps enforced before anything is sealed
// -----------------------------------------------------------------------

describe("table cap: refused before rendering, on every render (not just export)", () => {
  afterEach(() => uninstallFakeHost());

  it("an anchor that grew past 60 rows since it was linked refuses a later push, and never calls the relay", async () => {
    const links = await import("../../src/excel/links");
    const { helpers } = installFakeHost({ sheets: ["Model"] });
    const row = (r: number): string[] =>
      Array.from({ length: 3 }, (_col, c) => `${String(r)}x${String(c)}`);

    // 60 rows fits the cap: the export succeeds and anchors a hidden name.
    helpers.seed(
      "Model!A1:C60",
      Array.from({ length: 60 }, (_row, r) => row(r)),
    );
    helpers.select("Model!A1:C60");
    const relay = new FakeRelay();
    const ws = await createWorkspace(memoryStore());
    const { id } = await links.exportSelectionAsTable(ws, relay);

    // Excel rewrites a hidden name's formula when rows move under it
    // (link.local-transport-e2e.hunt.test.ts's "a moved source" pins the
    // same mechanic for a range link): the anchor now resolves to 61 rows,
    // one past TABLE_MAX_ROWS, with no new export in between.
    helpers.setNameFormula(anchorName(id), "=Model!$A$1:$C$61");
    helpers.seed(
      "Model!A1:C61",
      Array.from({ length: 61 }, (_row, r) => row(r)),
    );

    const putLinkSpy = vi.spyOn(relay, "putLink");
    const summary = await links.pushLinks("all", relay);

    expect(summary).toMatchObject({ pushed: 0, missing: 0, failed: 1 });
    expect(summary.failures[0]).toContain(TABLE_TOO_BIG);
    expect(putLinkSpy).not.toHaveBeenCalled();
  });
});

function memoryStore(): KeyStore {
  const map = new Map<string, string>();
  return {
    get: async (k) => map.get(k) ?? null,
    set: async (k, v) => {
      map.set(k, v);
    },
    remove: async (k) => {
      map.delete(k);
    },
  };
}
