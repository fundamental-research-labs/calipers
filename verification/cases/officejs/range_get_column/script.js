// officejs: write via Range.getColumn.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C3").getColumn(1).values = [[2], [5], [8]];
  await context.sync();
});
