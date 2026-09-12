// officejs: write via getIntersection.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C3").getIntersection(sheet.getRange("B2:D4")).values = [
    [1, 2],
    [3, 4],
  ];
  await context.sync();
});
