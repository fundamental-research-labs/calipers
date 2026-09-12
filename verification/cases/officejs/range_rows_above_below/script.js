// officejs: write via getRowsAbove and getRowsBelow.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A3:B3").values = [["x", "y"]];
  sheet.getRange("A3:B3").getRowsAbove(2).values = [
    [1, 2],
    [3, 4],
  ];
  sheet.getRange("A3:B3").getRowsBelow(1).values = [[5, 6]];
  await context.sync();
});
