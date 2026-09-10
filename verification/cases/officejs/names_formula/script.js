// officejs: workbook named formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [[10]];
  context.workbook.names.add("DoubleA1", "=2*Sheet1!$A$1");
  await context.sync();
});
