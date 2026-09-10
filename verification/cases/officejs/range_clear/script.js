// officejs: write values then clear the range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B2").values = [[1, 2], [3, 4]];
  sheet.getRange("A1:B2").clear();
  sheet.getRange("C1").values = [["cleared"]];
  await context.sync();
});
