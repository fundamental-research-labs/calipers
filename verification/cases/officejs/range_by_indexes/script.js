// officejs: write a 2x2 block via getRangeByIndexes.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRangeByIndexes(0, 0, 2, 2).values = [
    [1, 2],
    [3, 4],
  ];
  await context.sync();
});
