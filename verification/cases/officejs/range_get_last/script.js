// officejs: write via getLastCell / getLastRow / getLastColumn.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const block = sheet.getRange("A1:C3");
  block.values = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
  ];
  block.getLastCell().values = [["LC"]];
  block.getLastRow().values = [["a", "b", "c"]];
  block.getLastColumn().values = [[1], [2], [3]];
  await context.sync();
});
