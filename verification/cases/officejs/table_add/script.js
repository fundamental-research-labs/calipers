// officejs: create an Excel table from a small range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
  ];
  sheet.tables.add("A1:C4", true);
  await context.sync();
});
