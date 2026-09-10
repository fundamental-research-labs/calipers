// officejs: tiny workbook — two inputs and a sum.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [[10]];
  sheet.getRange("B1").values = [[20]];
  sheet.getRange("C1").formulas = [["=A1+B1"]];
  await context.sync();
});
