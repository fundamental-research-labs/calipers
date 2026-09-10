// officejs: 200x12 numeric grid write.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const rows = [];
  for (let r = 1; r <= 200; r++) {
    const row = [];
    for (let c = 1; c <= 12; c++) {
      row.push(r * 100 + c);
    }
    rows.push(row);
  }
  sheet.getRange("A1:L200").values = rows;
  await context.sync();
});
