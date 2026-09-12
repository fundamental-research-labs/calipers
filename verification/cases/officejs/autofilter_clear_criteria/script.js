// officejs: apply AutoFilter then clearCriteria.
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
  sheet.autoFilter.clearCriteria();
  await context.sync();
});
