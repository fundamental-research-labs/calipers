// officejs: write via Range.getRow.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C3").getRow(1).values = [[10, 20, 30]];
  await context.sync();
});
