// officejs: hide a row and a column.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C3").values = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
  ];
  sheet.getRange("A2").getEntireRow().rowHidden = true;
  sheet.getRange("B1").getEntireColumn().columnHidden = true;
  await context.sync();
});
