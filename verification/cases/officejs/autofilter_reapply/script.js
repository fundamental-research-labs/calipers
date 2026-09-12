// officejs: change a filtered cell then reapply AutoFilter.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B5").values = [
    ["Name", "Value"],
    ["a", 5],
    ["b", 15],
    ["c", 25],
    ["d", 8],
  ];
  sheet.autoFilter.apply(sheet.getRange("A1:B5"), 1, {
    criterion1: ">10",
    filterOn: Excel.FilterOn.custom,
  });
  sheet.getRange("B2").values = [[50]];
  sheet.autoFilter.reapply();
  await context.sync();
});
