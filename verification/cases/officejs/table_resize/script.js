// officejs: resize a table to include another row.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
  ];
  const table = sheet.tables.add("A1:C4", true);
  sheet.getRange("A5:C5").values = [["East", "C", 50]];
  table.resize("A1:C5");
  await context.sync();
});
