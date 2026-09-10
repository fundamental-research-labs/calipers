// officejs: AutoFilter on a data range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B4").values = [
    ["Name", "Value"],
    ["a", 1],
    ["b", 2],
    ["c", 3],
  ];
  sheet.autoFilter.apply(sheet.getRange("A1:B4"));
  await context.sync();
});
