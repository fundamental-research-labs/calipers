// officejs: 3-color scale conditional formatting.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A5").values = [[1], [5], [10], [15], [20]];
  const cf = sheet.getRange("A1:A5").conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
  cf.colorScale.criteria = {
    minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
    midpoint: { formula: "50", type: Excel.ConditionalFormatColorCriterionType.percentile, color: "#FFEB84" },
    maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" },
  };
  await context.sync();
});
