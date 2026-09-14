// officejs: AutoFilter.getRange and getRangeOrNullObject.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B4").values = [
    ["Name", "Value"],
    ["a", 1],
    ["b", 2],
    ["c", 3],
  ];
  sheet.autoFilter.apply(sheet.getRange("A1:B4"));
  const filtered = sheet.autoFilter.getRange();
  filtered.getCell(0, 0).values = [["Name"]];
  sheet.autoFilter.remove();
  const maybe = sheet.autoFilter.getRangeOrNullObject();
  await context.sync();
  sheet.getRange("D1").values = [[maybe.isNullObject ? "off" : "on"]];
  await context.sync();
});
