// officejs: 400 rows of values plus dependent formulas.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const vals = [];
  const fms = [];
  for (let i = 1; i <= 400; i++) {
    vals.push([i]);
    fms.push(["=A" + i + "*2"]);
  }
  sheet.getRange("A1:A400").values = vals;
  sheet.getRange("B1:B400").formulas = fms;
  await context.sync();
});
