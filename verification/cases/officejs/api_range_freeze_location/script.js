// Range query or formatting method with an observable result.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  sheet.getRange("B2:C3").values = [["alpha", "beta"], ["ALPHA", "alphabet"]];
  sheet.getRange("D5").format.fill.color = "#FFFF00";
  let output;
  sheet.freezePanes.freezeRows(2); const result = sheet.freezePanes.getLocation(); result.load("address"); await context.sync(); output = result.address;
  sheet.getRange("F1").values = [[output]];
  await context.sync();
});
