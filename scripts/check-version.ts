// package.json is the only place a version is typed by hand. Fails when the
// server crate, the manifests or the footer wiring disagree.
import { readFileSync } from "node:fs";

const root = new URL("../", import.meta.url);
const read = (file: string) => readFileSync(new URL(file, root), "utf8");
const version = (JSON.parse(read("package.json")) as { version: string }).version;

const problems: string[] = [];
if (!read("server/Cargo.toml").includes(`version = "${version}"`)) problems.push(`server/Cargo.toml is not ${version}`);
for (const file of ["manifest.xml", "manifest.prod.xml"]) {
  if (!read(file).includes(`<Version>${version}.0</Version>`)) problems.push(`${file} is not ${version}.0`);
}
if (read("taskpane.html").match(/v\d+\.\d+\.\d{3}/)) problems.push("taskpane.html hard-codes a version; use #app-version");

if (problems.length) {
  process.stderr.write(problems.join("\n") + "\n");
  process.exit(1);
}
process.stdout.write(`version ${version} consistent\n`);
