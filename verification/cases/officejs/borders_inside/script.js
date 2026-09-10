// officejs: inside horizontal/vertical borders.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B2").values = [[1, 2], [3, 4]];
  const b = sheet.getRange("A1:B2").format.borders;
  b.getItem("InsideHorizontal").style = Excel.BorderLineStyle.continuous;
  b.getItem("InsideVertical").style = Excel.BorderLineStyle.continuous;
  b.getItem("InsideHorizontal").color = "#666666";
  b.getItem("InsideVertical").color = "#666666";
  await context.sync();
});
