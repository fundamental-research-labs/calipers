package main

import "fmt"

// Each previously unsupported Excel.Functions method has a standalone case.
// Results and worksheet errors are persisted for manual Excel golden capture.
func functionSpecs() []spec {
	cases := []struct{ name, first, second string }{
		{"abs", `-12.5`, `0`},
		{"acos", `0.5`, `2`},
		{"acosh", `2`, `0`},
		{"asin", `0.5`, `2`},
		{"asinh", `2`, `-2`},
		{"atan", `1`, `-1`},
		{"atan2", `-1, 1`, `0, 0`},
		{"atanh", `0.5`, `1`},
		{"cos", `1`, `0`},
		{"cosh", `1`, `0`},
		{"sin", `1`, `0`},
		{"sinh", `1`, `-1`},
		{"tan", `1`, `0`},
		{"tanh", `1`, `-1`},
		{"degrees", `Math.PI`, `-Math.PI / 2`},
		{"radians", `180`, `-90`},
		{"exp", `1`, `0`},
		{"ln", `Math.E`, `0`},
		{"log", `8, 2`, `100`},
		{"log10", `1000`, `-1`},
		{"pi", ``, ``},
		{"power", `2, 10`, `-2, 3`},
		{"sqrt", `81`, `-1`},
		{"sqrtPi", `2`, `-1`},
		{"sign", `-8`, `0`},
		{"int", `-2.3`, `2.9`},
		{"trunc", `-12.345, 2`, `-12.345`},
		{"round", `-12.345, 2`, `125, -1`},
		{"roundUp", `-12.341, 2`, `121, -1`},
		{"roundDown", `-12.349, 2`, `129, -1`},
		{"mod", `-7, 3`, `7, 0`},
		{"quotient", `-7, 3`, `7, 0`},
		{"product", `sheet.getRange("A1:A3"), 2`, `3, 0, 4`},
		{"sum", `sheet.getRange("A1:A3"), f.abs(-4)`, `{ address: "Sheet1!A1:A3" }`},
		{"sumSq", `sheet.getRange("A1:A3")`, `-2, 3`},
		{"average", `sheet.getRange("A1:A3")`, `2, 4, 9`},
		{"min", `sheet.getRange("A1:A3"), -4`, `0, 2`},
		{"max", `sheet.getRange("A1:A3"), 4`, `-3, -2`},
		{"median", `sheet.getRange("A1:A3")`, `1, 2, 8, 9`},
		{"count", `sheet.getRange("A1:B3")`, `2, 4, 9`},
		{"countA", `sheet.getRange("A1:B3")`, `"", false, 0`},
		{"len", `"say \"hi\""`, `""`},
		{"left", `"abcdef", 3`, `"abcdef"`},
		{"right", `"abcdef", 3`, `"abcdef"`},
		{"mid", `"abcdef", 2, 3`, `"abcdef", 0, 2`},
		{"lower", `"HeLLo"`, `"A1 B2"`},
		{"upper", `"HeLLo"`, `"a1 b2"`},
		{"trim", `"  one   two  "`, `"   "`},
		{"concatenate", `"say ", "\"hi\"", 2`, `"=", "1+1"`},
		{"exact", `"Hello", "Hello"`, `"Hello", "hello"`},
	}
	var out []spec
	for _, c := range cases {
		out = append(out, sheetJS("api_function_"+c.name, "Excel.Functions."+c.name+": values and boundary behavior.", fmt.Sprintf(`
  sheet.name = "Sheet1";
  sheet.getRange("A1:B3").values = [[1, "text"], [2, true], [3, null]];
  const f = context.workbook.functions;
  const first = f.%s(%s);
  const second = f.%s(%s);
  first.load("value,error");
  second.load("value,error");
  await context.sync();
  // Preserve text beginning with '=' as text when writing through Range.values.
  const cellValue = result => result.error ? "" :
    (typeof result.value === "string" && /^[=+\-']/.test(result.value) ? "'" + result.value : result.value);
  sheet.getRange("D1:F3").values = [
    ["Scenario", "Value", "Error"],
    ["primary", cellValue(first), first.error || ""],
    ["boundary", cellValue(second), second.error || ""]
  ];
`, c.name, c.first, c.name, c.second)))
	}
	return out
}
