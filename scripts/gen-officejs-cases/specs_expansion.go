package main

// Keep each newly exposed method independently runnable on desktop Excel.
func expansionSpecs() []spec {
	var out []spec
	filters := []struct{ name, call string }{
		{"apply", `filter.apply({ filterOn: "Custom", criterion1: ">20" });`},
		{"values", `filter.applyValuesFilter(["20", "40"]);`},
		{"custom", `filter.applyCustomFilter(">10", "<40", "And");`},
		{"dynamic", `filter.applyDynamicFilter("AboveAverage");`},
		{"cell_color", `filter.applyCellColorFilter("#FFFF00");`},
		{"font_color", `filter.applyFontColorFilter("#FF0000");`},
		{"top_items", `filter.applyTopItemsFilter(2);`},
		{"bottom_items", `filter.applyBottomItemsFilter(2);`},
		{"top_percent", `filter.applyTopPercentFilter(50);`},
		{"bottom_percent", `filter.applyBottomPercentFilter(50);`},
		{"clear", `filter.applyValuesFilter(["20"]); filter.clear();`},
		{"table_clear", `filter.applyValuesFilter(["20"]); table.clearFilters();`},
		{"table_reapply", `filter.applyCustomFilter(">20"); await context.sync(); sheet.getRange("B2").values = [[50]]; table.reapplyFilters();`},
	}
	for _, f := range filters {
		out = append(out, sheetJS("api_filter_"+f.name, "Table column filtering and observable row visibility.", `
  sheet.getRange("A1:B5").values = [["Name", "Amount"], ["a", 10], ["b", 20], ["c", 30], ["d", 40]];
  sheet.getRange("B3").format.fill.color = "#FFFF00";
  sheet.getRange("B4").format.font.color = "#FF0000";
  const table = sheet.tables.add("A1:B5", true);
  const filter = table.columns.getItem("Amount").filter;
  `+f.call+`
  const rows = [2, 3, 4, 5].map(i => sheet.getRange("A" + i));
  rows.forEach(row => row.load("rowHidden"));
  await context.sync();
  sheet.getRange("D1:E5").values = [["Row", "Hidden"], ...rows.map((row, i) => [i + 2, row.rowHidden])];
`))
	}
	return out
}

func commentExpansionSpecs() []spec {
	cases := []struct{ name, call string }{
		{"count", `const result = comments.getCount(); await context.sync(); output = result.value;`},
		{"item", `first.load("id"); await context.sync(); const found = comments.getItem(first.id); found.load("content"); await context.sync(); output = found.content;`},
		{"at", `const found = comments.getItemAt(0); found.load("content"); await context.sync(); output = found.content;`},
		{"null", `const found = comments.getItemOrNullObject("missing-comment"); await context.sync(); output = found.isNullObject;`},
		{"cell", `const found = comments.getItemByCell(sheet.getRange("B2")); found.load("content"); await context.sync(); output = found.content;`},
		{"reply_id", `const reply = first.replies.add("reply"); reply.load("id"); await context.sync(); const found = comments.getItemByReplyId(reply.id); found.load("content"); await context.sync(); output = found.content;`},
		{"delete", `first.delete(); const result = comments.getCount(); await context.sync(); output = result.value;`},
		{"location", `const location = first.getLocation(); location.load("address"); await context.sync(); output = location.address;`},
		{"reply_add", `const reply = first.replies.add("reply"); reply.load("content"); await context.sync(); output = reply.content;`},
		{"reply_count", `first.replies.add("reply"); const result = first.replies.getCount(); await context.sync(); output = result.value;`},
		{"reply_item", `const reply = first.replies.add("reply"); reply.load("id"); await context.sync(); const found = first.replies.getItem(reply.id); found.load("content"); await context.sync(); output = found.content;`},
		{"reply_at", `first.replies.add("reply"); const found = first.replies.getItemAt(0); found.load("content"); await context.sync(); output = found.content;`},
		{"reply_null", `const found = first.replies.getItemOrNullObject("missing-reply"); await context.sync(); output = found.isNullObject;`},
		{"reply_delete", `const reply = first.replies.add("reply"); reply.delete(); const result = first.replies.getCount(); await context.sync(); output = result.value;`},
		{"reply_location", `const reply = first.replies.add("reply"); const location = reply.getLocation(); location.load("address"); await context.sync(); output = location.address;`},
		{"reply_parent", `const reply = first.replies.add("reply"); const parent = reply.getParentComment(); parent.load("content"); await context.sync(); output = parent.content;`},
	}
	var out []spec
	for _, c := range cases {
		out = append(out, sheetJS("api_comment_"+c.name, "Threaded comment method result, recorded without machine-specific IDs or author metadata.", `
  sheet.getRange("B2").values = [["anchor"]];
  const comments = context.workbook.comments;
  const first = comments.add(sheet.getRange("B2"), "root");
  let output;
  `+c.call+`
  sheet.getRange("D1").values = [[output]];
  // Remove threads after recording results so author and timestamp metadata
  // do not make the workbook comparison depend on the Windows account.
  const remaining = comments.load("items");
  await context.sync();
  remaining.items.forEach(comment => comment.delete());
`))
	}
	return out
}

func conditionalExpansionSpecs() []spec {
	cases := []struct{ name, call string }{
		{"count", `const result = formats.getCount(); await context.sync(); output = result.value;`},
		{"item", `first.load("id"); await context.sync(); const found = formats.getItem(first.id); found.load("type"); await context.sync(); output = found.type;`},
		{"at", `const found = formats.getItemAt(0); found.load("type"); await context.sync(); output = found.type;`},
		{"null", `const found = formats.getItemOrNullObject("missing-format"); await context.sync(); output = found.isNullObject;`},
		{"clear", `formats.clearAll(); const result = formats.getCount(); await context.sync(); output = result.value;`},
		{"delete", `first.delete(); const result = formats.getCount(); await context.sync(); output = result.value;`},
		{"range", `const range = first.getRange(); range.load("address"); await context.sync(); output = range.address;`},
		{"range_null", `const range = first.getRangeOrNullObject(); range.load("address"); await context.sync(); output = range.address;`},
	}
	var out []spec
	for _, c := range cases {
		out = append(out, sheetJS("api_cf_"+c.name, "Conditional format collection and range queries.", `
  sheet.getRange("B2:B5").values = [[1], [5], [10], [20]];
  const formats = sheet.getRange("B2:B5").conditionalFormats;
  const first = formats.add("CellValue");
  first.cellValue.rule = {formula1: "5", operator: "GreaterThan"};
  first.cellValue.format.fill.color = "#FFFF00";
  let output;
  `+c.call+`
  sheet.getRange("D1").values = [[output]];
`))
	}
	return out
}

func rangeQueryExpansionSpecs() []spec {
	cases := []struct{ name, call string }{
		{"range_used", `const result = sheet.getRange("B2:D8").getUsedRange(); result.load("address"); await context.sync(); output = result.address;`},
		{"range_used_null", `const result = sheet.getRange("K10:L12").getUsedRangeOrNullObject(true); await context.sync(); output = result.isNullObject;`},
		{"sheet_used", `const result = sheet.getUsedRange(true); result.load("address"); await context.sync(); output = result.address;`},
		{"sheet_used_null", `const empty = context.workbook.worksheets.add("Empty"); const result = empty.getUsedRangeOrNullObject(); await context.sync(); output = result.isNullObject;`},
		{"intersection_null", `const result = sheet.getRange("B2:C3").getIntersectionOrNullObject("C3:D4"); result.load("address"); const missing = sheet.getRange("B2").getIntersectionOrNullObject("D5"); await context.sync(); output = result.address + ":" + missing.isNullObject;`},
		{"find", `const result = sheet.getRange("B2:D5").find("ALPHA", {completeMatch: true}); result.load("address"); await context.sync(); output = result.address; result.values = [["found"]];`},
		{"find_null", `const result = sheet.getRange("B2:D5").findOrNullObject("absent", {}); await context.sync(); output = result.isNullObject;`},
		{"range_replace", `const result = sheet.getRange("B2:D5").replaceAll("alpha", "replaced", {completeMatch: true}); await context.sync(); output = result.value;`},
		{"sheet_replace", `const result = sheet.replaceAll("alpha", "replaced", {matchCase: false}); await context.sync(); output = result.value;`},
		{"freeze_location", `sheet.freezePanes.freezeRows(2); const result = sheet.freezePanes.getLocation(); result.load("address"); await context.sync(); output = result.address;`},
		{"freeze_location_null", `sheet.freezePanes.unfreeze(); const result = sheet.freezePanes.getLocationOrNullObject(); await context.sync(); output = result.isNullObject;`},
		{"adjust_indent", `sheet.getRange("B2").format.indentLevel = 1; sheet.getRange("B3").format.indentLevel = 4; sheet.getRange("B2:B3").format.adjustIndent(-2); const a = sheet.getRange("B2").format.load("indentLevel"); const b = sheet.getRange("B3").format.load("indentLevel"); await context.sync(); output = a.indentLevel + ":" + b.indentLevel;`},
	}
	var out []spec
	for _, c := range cases {
		out = append(out, sheetJS("api_range_"+c.name, "Range query or formatting method with an observable result.", `
  sheet.getRange("B2:C3").values = [["alpha", "beta"], ["ALPHA", "alphabet"]];
  sheet.getRange("D5").format.fill.color = "#FFFF00";
  let output;
  `+c.call+`
  sheet.getRange("F1").values = [[output]];
`))
	}
	return out
}
