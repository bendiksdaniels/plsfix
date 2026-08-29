// Excel's number separators: an application setting (ExcelApi 1.11) the pane
// reads to say what a format like #,##0.0 will actually show, never writes.

import { hostSupports } from "./internal";

export interface ExcelSeparators {
  decimal: string;
  thousands: string;
}

export async function readSeparators(): Promise<ExcelSeparators | null> {
  if (!hostSupports("1.11")) return null;
  return Excel.run(async (context) => {
    const application = context.application;
    application.load("decimalSeparator,thousandsSeparator");
    await context.sync();
    return {
      decimal: application.decimalSeparator,
      thousands: application.thousandsSeparator,
    };
  });
}
