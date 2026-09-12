// officejs: convert a table back to a range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
  ];
  const table = sheet.tables.add("A1:C4", true);
  const converted = table.convertToRange();
  converted.getCell(0, 0).values = [["plain"]];
  await context.sync();
});
