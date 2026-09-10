// officejs: icon-set conditional formatting.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A5").values = [[1], [5], [10], [15], [20]];
  const cf = sheet.getRange("A1:A5").conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
  cf.iconSet.style = Excel.IconSet.threeTrafficLights1;
  await context.sync();
});
