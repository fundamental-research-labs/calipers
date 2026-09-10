// officejs: nested IF / AND formulas.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B3").values = [
    [1, 10],
    [2, 20],
    [3, 30],
  ];
  sheet.getRange("C1").formulas = [["=IF(AND(A1>0,B1>5),B1-A1,0)"]];
  sheet.getRange("C2").formulas = [["=IF(OR(A2=2,B2>100),\"ok\",\"no\")"]];
  await context.sync();
});
