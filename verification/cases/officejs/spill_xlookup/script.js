// officejs: XLOOKUP formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B4").values = [
    ["a", 10],
    ["b", 20],
    ["c", 30],
    ["d", 40],
  ];
  sheet.getRange("D1").values = [["c"]];
  sheet.getRange("E1").formulas = [["=XLOOKUP(D1,A1:A4,B1:B4)"]];
  await context.sync();
});
