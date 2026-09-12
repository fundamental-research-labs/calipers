// officejs: write via Worksheet.getCell and Range.getCell.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getCell(0, 0).values = [["ws"]];
  sheet.getRange("A1:C3").getCell(1, 1).values = [["rng"]];
  await context.sync();
});
