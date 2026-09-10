// officejs: write Excel serial dates and a date number format.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A3").values = [[44927], [44928], [44929]];
  sheet.getRange("A1:A3").numberFormat = [["yyyy-mm-dd"], ["yyyy-mm-dd"], ["yyyy-mm-dd"]];
  await context.sync();
});
