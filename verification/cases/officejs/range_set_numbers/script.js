// officejs: set numeric values on a small range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C2").values = [
    [1, 2.5, -3],
    [0, 100, 0.125],
  ];
  await context.sync();
});
