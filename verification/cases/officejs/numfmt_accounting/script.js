// officejs: accounting number format.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A2").values = [[1234.5], [-50]];
  sheet.getRange("A1:A2").numberFormat = [["_($* #,##0.00_)"], ["_($* #,##0.00_)"]];
  await context.sync();
});
