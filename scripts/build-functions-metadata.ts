// Writes public/functions.json from src/functions/metadata.ts, the registration
// file Office reads to publish =SMT.ROUND and =SMT.ROUNDSUM. --check: exit 1 if
// the committed file differs (CI gate). Written via temp + rename so a reader
// never sees a half-written file, the same way the manifests are.
import { readFileSync, renameSync, writeFileSync } from "node:fs";
import { functionsMetadata } from "../src/functions/metadata";

const check = process.argv.includes("--check");
const file = "public/functions.json";
const path = new URL(`../${file}`, import.meta.url);
const expected = functionsMetadata();

let current = "";
try {
  current = readFileSync(path, "utf8");
} catch {
  current = "";
}

if (current === expected) {
  process.stdout.write(`${file} up to date\n`);
} else if (check) {
  process.stderr.write(
    `${file} differs from src/functions/metadata.ts - run npm run functions:build\n`,
  );
  process.exit(1);
} else {
  const tmp = new URL(`../${file}.tmp`, import.meta.url);
  writeFileSync(tmp, expected);
  renameSync(tmp, path);
  process.stdout.write(`wrote ${file}\n`);
}
