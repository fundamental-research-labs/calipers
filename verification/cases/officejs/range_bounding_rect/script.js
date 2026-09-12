// officejs: write via getBoundingRect.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").getBoundingRect(sheet.getRange("C3")).values = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
  ];
  await context.sync();
});
