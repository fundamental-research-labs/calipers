// officejs: write via getColumnsBefore and getColumnsAfter.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("C1:C2").values = [[9], [9]];
  sheet.getRange("C1:C2").getColumnsBefore(2).values = [
    [1, 2],
    [3, 4],
  ];
  sheet.getRange("C1:C2").getColumnsAfter(1).values = [[5], [6]];
  await context.sync();
});
