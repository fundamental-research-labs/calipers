// Conditional format collection and range queries.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  sheet.getRange("B2:B5").values = [[1], [5], [10], [20]];
  const formats = sheet.getRange("B2:B5").conditionalFormats;
  const first = formats.add("CellValue");
  first.cellValue.rule = {formula1: "5", operator: "GreaterThan"};
  first.cellValue.format.fill.color = "#FFFF00";
  let output;
  const found = formats.getItemOrNullObject("missing-format"); await context.sync(); output = found.isNullObject;
  sheet.getRange("D1").values = [[output]];
  await context.sync();
});
