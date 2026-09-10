// officejs: preset-criteria conditional formatting.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A6").values = [[1], [2], [3], [10], [11], [12]];
  const cf = sheet.getRange("A1:A6").conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
  cf.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.aboveAverage };
  cf.preset.format.fill.color = "lightblue";
  await context.sync();
});
