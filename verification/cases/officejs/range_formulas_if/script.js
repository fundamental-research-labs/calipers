// officejs: IF formulas over a small range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A3").values = [[5], [15], [25]];
  sheet.getRange("B1:B3").formulas = [["=IF(A1>=10,\"hi\",\"lo\")"], ["=IF(A2>=10,\"hi\",\"lo\")"], ["=IF(A3>=10,\"hi\",\"lo\")"]];
  await context.sync();
});
