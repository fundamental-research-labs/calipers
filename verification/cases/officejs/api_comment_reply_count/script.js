// Threaded comment method result, recorded without machine-specific IDs or author metadata.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  sheet.getRange("B2").values = [["anchor"]];
  const comments = context.workbook.comments;
  const first = comments.add(sheet.getRange("B2"), "root");
  let output;
  first.replies.add("reply"); const result = first.replies.getCount(); await context.sync(); output = result.value;
  sheet.getRange("D1").values = [[output]];
  // Remove threads after recording results so author and timestamp metadata
  // do not make the workbook comparison depend on the Windows account.
  const remaining = comments.load("items");
  await context.sync();
  remaining.items.forEach(comment => comment.delete());
  await context.sync();
});
