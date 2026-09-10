// officejs: set boolean cell values.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A4").values = [[true], [false], [true], [false]];
  await context.sync();
});
