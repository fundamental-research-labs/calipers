// officejs: spilling VSTACK formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A2").values = [[1], [2]];
  sheet.getRange("C1:C2").values = [[3], [4]];
  sheet.getRange("E1").formulas = [["=VSTACK(A1:A2,C1:C2)"]];
  await context.sync();
});
