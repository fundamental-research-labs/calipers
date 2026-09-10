// officejs: spilling TRANSPOSE formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C1").values = [[1, 2, 3]];
  sheet.getRange("A3").formulas = [["=TRANSPOSE(A1:C1)"]];
  await context.sync();
});
