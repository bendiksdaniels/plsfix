// Pass-2 property: src/link/crypto.ts's seal/open round trip. A round trip
// with the same key and AAD returns the exact plaintext; a wrong key, a
// wrong AAD, or a single flipped byte in the blob never opens. Fixed seed,
// deterministic.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { open, seal } from "../../src/link/crypto";

const SEED = 20260927;
const ROUND_TRIP_RUNS = 200;
const TAMPER_RUNS = 150;

const keyArb = fc.uint8Array({ minLength: 32, maxLength: 32 });
const aadArb = fc.string({ maxLength: 40 });
const plaintextArb = fc.uint8Array({ minLength: 0, maxLength: 2000 });

describe("crypto seal/open round trip", () => {
  it("returns exactly the plaintext given to it, for any key, AAD and bytes", async () => {
    await fc.assert(
      fc.asyncProperty(
        keyArb,
        aadArb,
        plaintextArb,
        async (key, aad, plaintext) => {
          const blob = await seal(key, aad, plaintext);
          const opened = await open(key, aad, blob);
          expect(opened).toEqual(plaintext);
        },
      ),
      { seed: SEED, numRuns: ROUND_TRIP_RUNS },
    );
  });

  it("blob length is always 12-byte IV + plaintext + 16-byte tag", async () => {
    await fc.assert(
      fc.asyncProperty(
        keyArb,
        aadArb,
        plaintextArb,
        async (key, aad, plain) => {
          const blob = await seal(key, aad, plain);
          expect(blob.length).toBe(12 + plain.length + 16);
        },
      ),
      { seed: SEED, numRuns: ROUND_TRIP_RUNS },
    );
  });
});

describe("crypto seal/open: a wrong key or AAD never opens", () => {
  it("a different random key always fails to open, never a different plaintext", async () => {
    await fc.assert(
      fc.asyncProperty(
        keyArb,
        keyArb,
        aadArb,
        plaintextArb,
        async (keyA, keyB, aad, plaintext) => {
          fc.pre(!keyA.every((byte, index) => byte === keyB[index]));
          const blob = await seal(keyA, aad, plaintext);
          await expect(open(keyB, aad, blob)).rejects.toThrow(/decrypt/);
        },
      ),
      { seed: SEED, numRuns: TAMPER_RUNS },
    );
  });

  it("a different AAD string always fails to open", async () => {
    await fc.assert(
      fc.asyncProperty(
        keyArb,
        aadArb,
        aadArb,
        plaintextArb,
        async (key, aadA, aadB, plaintext) => {
          fc.pre(aadA !== aadB);
          const blob = await seal(key, aadA, plaintext);
          await expect(open(key, aadB, blob)).rejects.toThrow(/decrypt/);
        },
      ),
      { seed: SEED, numRuns: TAMPER_RUNS },
    );
  });
});

describe("crypto seal/open: a flipped byte never opens", () => {
  it("flipping any single byte of the blob (IV, ciphertext or tag) always fails to open", async () => {
    await fc.assert(
      fc.asyncProperty(
        keyArb,
        aadArb,
        plaintextArb,
        fc.nat(),
        async (key, aad, plaintext, flipRaw) => {
          const blob = await seal(key, aad, plaintext);
          const flipIndex = flipRaw % blob.length;
          const tampered = blob.slice();
          tampered[flipIndex] = (tampered[flipIndex]! ^ 0xff) & 0xff;
          await expect(open(key, aad, tampered)).rejects.toThrow(/decrypt/);
        },
      ),
      { seed: SEED, numRuns: TAMPER_RUNS },
    );
  });

  it("truncating the blob below IV+tag length always fails to open, never throws past the stage", async () => {
    await fc.assert(
      fc.asyncProperty(
        keyArb,
        aadArb,
        fc.integer({ min: 0, max: 27 }),
        async (key, aad, shortLength) => {
          const short = new Uint8Array(shortLength);
          await expect(open(key, aad, short)).rejects.toThrow(/decrypt/);
        },
      ),
      { seed: SEED, numRuns: 28 },
    );
  });
});
