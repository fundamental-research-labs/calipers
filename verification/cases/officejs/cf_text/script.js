// officejs: contains-text conditional formatting.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A4").values = [["East"], ["West"], ["East"], ["North"]];
  const cf = sheet.getRange("A1:A4").conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  cf.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "East" };
  cf.textComparison.format.font.bold = true;
  await context.sync();
});
