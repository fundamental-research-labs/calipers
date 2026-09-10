// officejs: add a column to a table.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B3").values = [
    ["Name", "Value"],
    ["a", 1],
    ["b", 2],
  ];
  const table = sheet.tables.add("A1:B3", true);
  table.columns.add(null, [["Extra"], [10], [20]]);
  await context.sync();
});
