// officejs: spilling SORT formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A6").values = [[5], [1], [4], [2], [3], [0]];
  sheet.getRange("C1").formulas = [["=SORT(A1:A6,1,1)"]];
  await context.sync();
});
