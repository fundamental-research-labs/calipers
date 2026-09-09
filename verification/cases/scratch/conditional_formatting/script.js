// Scratch: add cell-value conditional formatting.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A5").values = [[1], [5], [10], [15], [20]];
  const cf = sheet.getRange("A1:A5").conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
  cf.cellValue.rule = {
    formula1: "10",
    operator: Excel.ConditionalCellValueOperator.greaterThan,
  };
  cf.cellValue.format.fill.color = "yellow";
  await context.sync();
});
