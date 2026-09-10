// officejs: spilling FILTER formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B6").values = [
    ["Region", "Sales"],
    ["East", 10],
    ["West", 20],
    ["East", 30],
    ["West", 40],
    ["East", 15],
  ];
  sheet.getRange("D1").formulas = [["=FILTER(A2:B6,A2:A6=\"East\")"]];
  await context.sync();
});
