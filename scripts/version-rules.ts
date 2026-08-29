// What "the version is consistent" means, as rules over file text: the pure
// half of scripts/check-version.ts, so the gate itself can be tested without a
// repo on disk. `read` is injected; nothing here touches the filesystem or the
// process.

// A version anywhere in a pane or its boot code is a version that will go
// stale, so any vMAJOR.MINOR.PATCH shape counts. The patch is `\d+`, not
// `\d{3}`: `npm version patch` writes 2.1.18, so a three-digit rule could only
// ever fire on the zero-padded form the footer renders (v2.1.018) and never on
// the plain string a person would paste.
const HARD_CODED = /v\d+\.\d+\.\d+/;

// The same rule for TypeScript, where a version reaches a pane as a string
// literal: `"v2.` is a hard-coded version whether or not the patch is written
// out. A template that interpolates APP_VERSION is not one.
const HARD_CODED_LITERAL = /["'`]v\d+\./;

// Every file the gate reads and what has to be true of it, in the order the
// messages should read.
export function versionProblems(
  read: (file: string) => string,
  version: string,
): string[] {
  const problems: string[] = [];
  // Only the [package] table counts: a dependency pinned at the same string
  // would otherwise satisfy the gate over a stale crate version, which is why
  // scripts/bump-patch.sh anchors its rewrite to that table too.
  const crate = read("server/Cargo.toml")
    .split(/^\[/m)
    .find((table) => table.startsWith("package]"));
  if (crate === undefined || !crate.includes(`version = "${version}"`))
    problems.push(`server/Cargo.toml is not ${version}`);
  for (const file of ["manifest.xml", "manifest.prod.xml"]) {
    if (!read(file).includes(`<Version>${version}.0</Version>`))
      problems.push(`${file} is not ${version}.0`);
  }
  // Both panes show the version through #app-version, and both boot files put
  // it there from __APP_VERSION__; any of the four could hard-code it and go
  // stale, so the gate looks at all four.
  for (const page of ["taskpane.html", "pptpane.html"]) {
    if (HARD_CODED.test(read(page)))
      problems.push(`${page} hard-codes a version; use #app-version`);
  }
  for (const module of ["src/main.ts", "src/ppt/main.ts"]) {
    if (HARD_CODED_LITERAL.test(read(module)))
      problems.push(`${module} hard-codes a version; use APP_VERSION`);
  }
  return problems;
}
