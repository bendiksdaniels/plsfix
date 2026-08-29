// The brand palette that travels with the file: the same JSON the pane keeps in
// localStorage, held in the workbook setting "plsfix.brand.v1". Precedence: the
// workbook setting wins over localStorage, and localStorage is only the default
// a new workbook starts from - so a model keeps its brand on another computer.

const BRAND_SETTING = "plsfix.brand.v1";

// Absent and blank both mean "this workbook has no palette of its own", so the
// caller keeps the machine default rather than dropping to the shipped colors.
// The JSON is handed back verbatim: parsing and rejecting it is settings.ts's
// job, and a string this file cannot read is not a string it may erase.
export async function readWorkbookBrand(): Promise<string | null> {
  return Excel.run(async (context) => {
    const setting =
      context.workbook.settings.getItemOrNullObject(BRAND_SETTING);
    setting.load("isNullObject,value");
    await context.sync();
    if (setting.isNullObject) return null;

    const json = String(setting.value);
    return json.trim() === "" ? null : json;
  });
}

// One setting per workbook: add() replaces the value under the same key, so the
// file carries the palette the modeller last chose and nothing older.
export async function writeWorkbookBrand(json: string): Promise<void> {
  await Excel.run(async (context) => {
    context.workbook.settings.add(BRAND_SETTING, json);
    await context.sync();
  });
}
