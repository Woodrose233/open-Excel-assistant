/* global Office */

export async function insertText(text: string) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.insert(Excel.InsertShiftDirection.down);
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
