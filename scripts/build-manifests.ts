// Writes manifest.xml (dev) and manifest.prod.xml (prod) from manifest/spec.ts.
// --check: exit 1 if a committed file differs (CI gate). Files are written via
// temp + rename so a hard-linked copy elsewhere never sees a half-written file.
import { readFileSync, renameSync, writeFileSync } from "node:fs";
import { ADDIN, ENVIRONMENTS } from "../manifest/spec";
import { buildManifest } from "../manifest/xml";

const check = process.argv.includes("--check");
let drift = 0;
for (const env of ENVIRONMENTS) {
  const expected = buildManifest(env, ADDIN);
  const path = new URL(`../${env.file}`, import.meta.url);
  let current = "";
  try {
    current = readFileSync(path, "utf8");
  } catch {
    current = "";
  }
  if (current === expected) continue;
  if (check) {
    process.stderr.write(
      `${env.file} differs from manifest/spec.ts - run npm run manifest:build\n`,
    );
    drift += 1;
    continue;
  }
  const tmp = new URL(`../${env.file}.tmp`, import.meta.url);
  writeFileSync(tmp, expected);
  renameSync(tmp, path);
  process.stdout.write(`wrote ${env.file}\n`);
}
process.exit(drift ? 1 : 0);
