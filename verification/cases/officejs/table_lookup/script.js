// officejs: table collection getCount / getItem / getItemAt / getItemOrNullObject and Table.getRange.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
  ];
  const created = sheet.tables.add("A1:C4", true);
  created.name = "Sales";
  const byName = sheet.tables.getItem("Sales");
  const full = byName.getRange();
  full.getCell(0, 0).values = [["Region"]];
  const byIndex = sheet.tables.getItemAt(0);
  byIndex.name = "SalesTable";
  const missing = sheet.tables.getItemOrNullObject("Nope");
  const count = sheet.tables.getCount();
  await context.sync();
  sheet.getRange("E1").values = [[count.value]];
  sheet.getRange("F1").values = [[missing.isNullObject ? "gone" : "hit"]];
  await context.sync();
});
