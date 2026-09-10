// officejs: spilling UNIQUE formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A8").values = [["a"], ["b"], ["a"], ["c"], ["b"], ["d"], ["a"], ["e"]];
  sheet.getRange("C1").formulas = [["=UNIQUE(A1:A8)"]];
  await context.sync();
});
