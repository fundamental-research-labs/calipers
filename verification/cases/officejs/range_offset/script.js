// officejs: write via getOffsetRange from A1.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [[1]];
  sheet.getRange("A1").getOffsetRange(1, 1).values = [[2]];
  sheet.getRange("A1").getOffsetRange(2, 2).values = [[3]];
  await context.sync();
});
