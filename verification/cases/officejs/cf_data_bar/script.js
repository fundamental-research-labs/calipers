// officejs: data-bar conditional formatting.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A5").values = [[1], [5], [10], [15], [20]];
  sheet.getRange("A1:A5").conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
  await context.sync();
});
