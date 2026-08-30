// The brand colour rules every chart follows, on the sheet (src/excel/charts.ts)
// and on the slide (src/excel/link-chart.ts decides, src/ppt draws): the
// series palette and its order, the waterfall's totals, rises and falls, and a
// pie's cycle through the palette. Pure: brand colours in, hex strings out.

import { bridgeSeries } from "./chartmath";
import { tint } from "./settings";

export interface SeriesBrand {
  primary: string;
  accent: string;
}

export interface WaterfallBrand extends SeriesBrand {
  external: string;
}

const PRIMARY_TINT = 0.55;
const ACCENT_TINT = 0.45;
const PRIMARY_LIGHT_TINT = 0.78;
const ACCENT_LIGHT_TINT = 0.7;
// bridgeSeries needs an opening, one step and a closing; fewer is all totals.
const MIN_BRIDGE = 3;

// Six colours: the accent first, the primary second, then their tints, which
// is the order "Brand-format chart" has always painted series in.
export function seriesPalette({ primary, accent }: SeriesBrand): string[] {
  return [
    accent,
    primary,
    tint(primary, PRIMARY_TINT),
    tint(accent, ACCENT_TINT),
    tint(primary, PRIMARY_LIGHT_TINT),
    tint(accent, ACCENT_LIGHT_TINT),
  ];
}

// The first and last points are the totals, branded primary; between them a
// fall wears the external colour and a rise the accent. Office.js has no "set
// as total" flag for waterfall points, so position is the rule on the sheet
// and on the slide alike.
export function waterfallColors(
  values: number[],
  brand: WaterfallBrand,
): string[] {
  // Two points are an opening and a closing total with nothing between them.
  if (values.length < MIN_BRIDGE) return values.map(() => brand.primary);
  const bridge = bridgeSeries(values);
  const last = values.length - 1;
  return values.map((_value, index) => {
    if (index === 0 || index === last) return brand.primary;
    return (bridge.fall[index] ?? 0) > 0 ? brand.external : brand.accent;
  });
}

// One colour per slice, cycling the palette past its sixth entry.
export function pieColors(count: number, brand: SeriesBrand): string[] {
  const palette = seriesPalette(brand);
  return Array.from(
    { length: count },
    (_, index) => palette[index % palette.length]!,
  );
}
