// Generates the three PNG sizes Office requires from a small, semantic SVG
// icon set. The source is deliberately vector-first: crisp 16px commands are
// more useful in a ribbon than illustrative artwork.
import { mkdirSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const output = fileURLToPath(
  new URL("../public/assets/ribbon/", import.meta.url),
);
const navy = "#14213D";
const teal = "#2EC4B6";
const ink = `fill="none" stroke="${navy}" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"`;
const accent = `fill="none" stroke="${teal}" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"`;

const icons = {
  pane: `<rect x="12" y="10" width="40" height="44" rx="5" ${ink}/><path d="M22 20h20M22 30h20M22 40h12" ${ink}/><path d="M45 39v8M41 43h8" ${accent}/>`,
  palette: `<path d="M32 10C19 10 10 19 10 31s9 22 21 22h4c3 0 4-3 2-5-2-3 1-7 5-7h2c6 0 10-4 10-10 0-12-10-21-22-21Z" ${ink}/><circle cx="21" cy="29" r="3" fill="${teal}"/><circle cx="31" cy="20" r="3" fill="${teal}"/><circle cx="42" cy="27" r="3" fill="${teal}"/>`,
  "fill-right": `<path d="M12 17h24M12 32h24M12 47h24M19 10v44M36 10v44" ${ink}/><path d="M41 32h12M48 25l7 7-7 7" ${accent}/>`,
  "fill-down": `<path d="M17 12v24M32 12v24M47 12v24M10 19h44M10 36h44" ${ink}/><path d="M32 41v12M25 48l7 7 7-7" ${accent}/>`,
  guard: `<path d="M32 9l18 7v13c0 12-8 21-18 26-10-5-18-14-18-26V16l18-7Z" ${ink}/><path d="m23 32 6 6 12-14" ${accent}/>`,
  audit: `<circle cx="28" cy="28" r="14" ${ink}/><path d="m39 39 13 13M22 28h12M28 22v12" ${accent}/>`,
  precedents: `<circle cx="49" cy="32" r="4" fill="${navy}"/><circle cx="17" cy="18" r="4" fill="${teal}"/><circle cx="17" cy="46" r="4" fill="${teal}"/><path d="M22 18h8c8 0 8 14 15 14h4M22 46h8c8 0 8-14 15-14h4" ${ink}/><path d="m24 12-7 6 7 6M24 40l-7 6 7 6" ${accent}/>`,
  dependents: `<circle cx="15" cy="32" r="4" fill="${navy}"/><circle cx="47" cy="18" r="4" fill="${teal}"/><circle cx="47" cy="46" r="4" fill="${teal}"/><path d="M19 32h8c7 0 7-14 15-14h5M19 32h8c7 0 7 14 15 14h5" ${ink}/><path d="m40 12 7 6-7 6M40 40l7 6-7 6" ${accent}/>`,
  undo: `<path d="M25 19 13 31l12 12M15 31h20c9 0 16 5 16 14v2" ${ink}/><path d="M45 48h7" ${accent}/>`,
  "paste-values": `<rect x="18" y="13" width="29" height="40" rx="4" ${ink}/><path d="M25 13v-4h15v4M25 27h15M25 37h6M37 37h3M25 45h14" ${ink}/><path d="M30 37h4M38 37h3" ${accent}/>`,
  "paste-formats": `<rect x="14" y="16" width="25" height="33" rx="4" ${ink}/><path d="M20 16v-5h13v5M42 22l9 9-16 16-10 1 1-10 16-16Z" ${accent}/><path d="m39 25 9 9" ${accent}/>`,
  cagr: `<path d="M12 51V13M12 51h41M20 42l10-10 8 5 14-18" ${ink}/><path d="M43 19h9v9" ${accent}/>`,
  sign: `<path d="M13 21h16M21 13v16M37 21h15" ${ink}/><path d="M14 43h36" ${accent}/>`,
  "scale-up": `<path d="M32 52V13M22 23l10-10 10 10" ${accent}/><path d="M14 48h11M14 40h11M39 40h11M39 48h11" ${ink}/>`,
  "scale-down": `<path d="M32 12v39M22 41l10 10 10-10" ${accent}/><path d="M14 16h11M14 24h11M39 16h11M39 24h11" ${ink}/>`,
  waterfall: `<path d="M11 51h42M14 45h8V28h10v12h10V18h8v27" ${ink}/><path d="M22 28v17M32 40v5M42 18v27" ${accent}/>`,
  toc: `<rect x="14" y="10" width="36" height="44" rx="4" ${ink}/><path d="M23 21h18M23 32h18M23 43h12" ${ink}/><circle cx="19" cy="21" r="2" fill="${teal}"/><circle cx="19" cy="32" r="2" fill="${teal}"/><circle cx="19" cy="43" r="2" fill="${teal}"/>`,
  links: `<path d="m26 39-4 4c-5 5-12-2-7-7l10-10c5-5 12 2 7 7l-3 3M38 25l4-4c5-5 12 2 7 7L39 38c-5 5-12-2-7-7l3-3" ${ink}/><path d="m25 39 14-14" ${accent}/>`,
} as const;

function svg(body: string): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">${body}</svg>`;
}

mkdirSync(output, { recursive: true });
for (const [name, body] of Object.entries(icons)) {
  const source = `${output}${name}.svg`;
  writeFileSync(source, svg(body));
  for (const size of [16, 32, 80]) {
    // macOS's native renderer preserves SVG strokes reliably; ImageMagick's
    // installed SVG delegate drops them on this machine, leaving only fills.
    const result = spawnSync("sips", [
      source,
      "-z",
      String(size),
      String(size),
      "-s",
      "format",
      "png",
      "--out",
      `${output}${name}-${size}.png`,
    ]);
    if (result.status !== 0) {
      throw new Error(`Could not render ribbon icon ${name}: ${result.stderr}`);
    }
  }
}
