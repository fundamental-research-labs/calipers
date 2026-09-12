// officejs: workbook named formula via addFormulaLocal.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [[10]];
  context.workbook.names.addFormulaLocal("DoubleA1", "=2*Sheet1!$A$1");
  sheet.getRange("B1").formulas = [["=DoubleA1"]];
  await context.sync();
});
