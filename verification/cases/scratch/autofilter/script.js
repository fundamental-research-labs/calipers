// Scratch: apply an AutoFilter to a data range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B3").values = [
    ["Name", "Value"],
    ["a", 1],
    ["b", 2],
  ];
  sheet.autoFilter.apply(sheet.getRange("A1:B3"));
  await context.sync();
});
