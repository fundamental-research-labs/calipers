// officejs: INDEX/MATCH lookup formulas.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B4").values = [
    ["Key", "Val"],
    ["a", 10],
    ["b", 20],
    ["c", 30],
  ];
  sheet.getRange("D1").values = [["b"]];
  sheet.getRange("E1").formulas = [["=INDEX(B2:B4,MATCH(D1,A2:A4,0))"]];
  await context.sync();
});
