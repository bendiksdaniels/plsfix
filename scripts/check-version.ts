// package.json is the only place a version is typed by hand. Fails when the
// server crate, the manifests or the pane wiring disagree. This file is the
// I/O edge only: the rules live in version-rules.ts, where they are tested.
import { readFileSync } from "node:fs";
import { versionProblems } from "./version-rules";

const root = new URL("../", import.meta.url);
const read = (file: string) => readFileSync(new URL(file, root), "utf8");
const version = (JSON.parse(read("package.json")) as { version: string })
  .version;

const problems = versionProblems(read, version);
if (problems.length) {
  process.stderr.write(problems.join("\n") + "\n");
  process.exit(1);
}
process.stdout.write(`version ${version} consistent\n`);
