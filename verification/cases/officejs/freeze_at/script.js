// officejs: freeze panes at a cell.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C3").values = [
    ["Name", "Qty", "Price"],
    ["a", 1, 2],
    ["b", 3, 4],
  ];
  sheet.freezePanes.freezeAt(sheet.getRange("B2"));
  await context.sync();
});
