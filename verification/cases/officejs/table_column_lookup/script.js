// officejs: table column getItemAt and getItemOrNullObject.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
  ];
  const table = sheet.tables.add("A1:C4", true);
  const product = table.columns.getItemAt(1);
  product.name = "Prod";
  const missing = table.columns.getItemOrNullObject("Nope");
  await context.sync();
  sheet.getRange("E1").values = [[missing.isNullObject ? "gone" : "hit"]];
  await context.sync();
});
