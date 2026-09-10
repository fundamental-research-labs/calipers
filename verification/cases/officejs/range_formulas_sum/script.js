// officejs: SUM / AVERAGE formulas over a small range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A4").values = [[1], [2], [3], [4]];
  sheet.getRange("B1").formulas = [["=SUM(A1:A4)"]];
  sheet.getRange("B2").formulas = [["=AVERAGE(A1:A4)"]];
  await context.sync();
});
