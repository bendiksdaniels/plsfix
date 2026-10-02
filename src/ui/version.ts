// Version: formats package.json's semver for every place the pane shows its
// version (footer, error reports). Owns the one vMAJOR.MINOR.PATCH format -
// the patch is always zero-padded to three digits (v1.1.002). Anything that
// is not exactly MAJOR.MINOR.PATCH throws rather than rendering a partial string.
export function formatVersion(semver: string): string {
  const match = /^(\d+)\.(\d+)\.(\d+)$/.exec(semver);
  if (!match) throw new Error(`not a semver version: ${semver}`);
  return `v${match[1]}.${match[2]}.${match[3]!.padStart(3, "0")}`;
}
