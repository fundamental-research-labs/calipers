// Range query or formatting method with an observable result.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  sheet.getRange("B2:C3").values = [["alpha", "beta"], ["ALPHA", "alphabet"]];
  sheet.getRange("D5").format.fill.color = "#FFFF00";
  let output;
  const empty = context.workbook.worksheets.add("Empty"); const result = empty.getUsedRangeOrNullObject(); await context.sync(); output = result.isNullObject;
  sheet.getRange("F1").values = [[output]];
  await context.sync();
});
