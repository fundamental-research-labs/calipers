// officejs: three sheets, 200 rows each.
await Excel.run(async (context) => {
  const names = ["One", "Two", "Three"];
  for (let n = 0; n < names.length; n++) {
    const s = n === 0
      ? context.workbook.worksheets.getActiveWorksheet()
      : context.workbook.worksheets.add(names[n]);
    if (n === 0) {
      s.name = names[0];
    }
    await context.sync();
    const rows = [];
    for (let i = 1; i <= 200; i++) {
      rows.push([names[n], i, i * 2]);
    }
    s.getRange("A1:C200").values = rows;
    await context.sync();
  }
});
