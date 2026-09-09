// Scratch: write numeric values into an empty workbook.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B2").values = [
    [1, 2.5],
    [-3, 0],
  ];
  await context.sync();
});
