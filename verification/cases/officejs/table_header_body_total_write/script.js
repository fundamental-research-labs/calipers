// officejs: write header, body, and total table ranges.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
  ];
  const table = sheet.tables.add("A1:C4", true);
  table.showTotals = true;
  table.getHeaderRowRange().values = [["R", "P", "S"]];
  table.getDataBodyRange().values = [
    ["North", "X", 1],
    ["South", "Y", 2],
    ["West", "Z", 3],
  ];
  table.getTotalRowRange().values = [["Total", "", 6]];
  await context.sync();
});
