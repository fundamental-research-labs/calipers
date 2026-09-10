// officejs: outline borders on a block.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C3").values = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
  ];
  const b = sheet.getRange("A1:C3").format.borders;
  b.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
  b.getItem("EdgeBottom").style = Excel.BorderLineStyle.continuous;
  b.getItem("EdgeLeft").style = Excel.BorderLineStyle.continuous;
  b.getItem("EdgeRight").style = Excel.BorderLineStyle.continuous;
  await context.sync();
});
