// officejs: spilling SORTBY formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B5").values = [
    ["n", 3],
    ["a", 1],
    ["c", 4],
    ["b", 2],
    ["d", 0],
  ];
  sheet.getRange("D1").formulas = [["=SORTBY(A1:A5,B1:B5,1)"]];
  await context.sync();
});
