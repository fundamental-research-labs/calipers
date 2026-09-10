// officejs: copyFrom a source range to a destination.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A3").values = [[1], [2], [3]];
  sheet.getRange("C1").copyFrom(sheet.getRange("A1:A3"));
  await context.sync();
});
