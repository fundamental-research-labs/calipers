// officejs: write via getResizedRange and getAbsoluteResizedRange.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").getResizedRange(1, 1).values = [
    [1, 2],
    [3, 4],
  ];
  sheet.getRange("D1").getAbsoluteResizedRange(2, 2).values = [
    [5, 6],
    [7, 8],
  ];
  await context.sync();
});
