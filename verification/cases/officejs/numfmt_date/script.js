// officejs: date number format on serials.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A2").values = [[44927], [45300]];
  sheet.getRange("A1:A2").numberFormat = [["m/d/yyyy"], ["yyyy-mm-dd"]];
  await context.sync();
});
