// officejs: XMATCH formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A4").values = [["a"], ["b"], ["c"], ["d"]];
  sheet.getRange("C1").values = [["c"]];
  sheet.getRange("D1").formulas = [["=XMATCH(C1,A1:A4)"]];
  await context.sync();
});
