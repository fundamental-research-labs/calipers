// officejs: currency number format.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A3").values = [[1.5], [20], [-3.25]];
  sheet.getRange("A1:A3").numberFormat = [["$#,##0.00"], ["$#,##0.00"], ["$#,##0.00"]];
  await context.sync();
});
