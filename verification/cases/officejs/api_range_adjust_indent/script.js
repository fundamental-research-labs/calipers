// Range query or formatting method with an observable result.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  sheet.getRange("B2:C3").values = [["alpha", "beta"], ["ALPHA", "alphabet"]];
  sheet.getRange("D5").format.fill.color = "#FFFF00";
  let output;
  sheet.getRange("B2").format.indentLevel = 1; sheet.getRange("B3").format.indentLevel = 4; sheet.getRange("B2:B3").format.adjustIndent(-2); const a = sheet.getRange("B2").format.load("indentLevel"); const b = sheet.getRange("B3").format.load("indentLevel"); await context.sync(); output = a.indentLevel + ":" + b.indentLevel;
  sheet.getRange("F1").values = [[output]];
  await context.sync();
});
