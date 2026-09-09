// Scratch: create an Excel table.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B3").values = [
    ["Name", "Value"],
    ["a", 1],
    ["b", 2],
  ];
  sheet.tables.add("A1:B3", true);
  await context.sync();
});
