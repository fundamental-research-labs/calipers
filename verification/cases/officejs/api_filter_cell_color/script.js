// Table column filtering and observable row visibility.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  sheet.getRange("A1:B5").values = [["Name", "Amount"], ["a", 10], ["b", 20], ["c", 30], ["d", 40]];
  sheet.getRange("B3").format.fill.color = "#FFFF00";
  sheet.getRange("B4").format.font.color = "#FF0000";
  const table = sheet.tables.add("A1:B5", true);
  const filter = table.columns.getItem("Amount").filter;
  filter.applyCellColorFilter("#FFFF00");
  const rows = [2, 3, 4, 5].map(i => sheet.getRange("A" + i));
  rows.forEach(row => row.load("rowHidden"));
  await context.sync();
  sheet.getRange("D1:E5").values = [["Row", "Hidden"], ...rows.map((row, i) => [i + 2, row.rowHidden])];
  await context.sync();
});
