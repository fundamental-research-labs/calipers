package xlsxmodel

import (
	"strings"
	"testing"
)

func band(min, max float64) CellBand {
	return CellBand{Min: &min, Max: &max}
}

func TestCompareFillBgColorException(t *testing.T) {
	excel := solidFillXLSX(t, "indexed:64")
	mog := solidFillXLSX(t, "")
	opt := Options{Ignore: []string{IgnoreFillBgColor}}

	got, err := Compare(mog, excel, opt)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("indexed:64 vs omitted bg must be equal with %s, got %v", IgnoreFillBgColor, got.Diffs)
	}

	got, err = Compare(mog, excel)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "styles", "Sheet1!A1") {
		t.Fatalf("without exception want styles mismatch, got %v", got.Diffs)
	}

	changed := solidFillXLSX(t, "")
	// unrelated font change on the same yellow fill
	other := mustXLSX(t, map[string]string{
		"xl/styles.xml":            solidFillStyles("indexed:64", `<fonts count="1"><font><sz val="20"/><name val="Calibri"/></font></fonts>`),
		"xl/worksheets/sheet1.xml": solidFillSheet(),
	})
	got, err = Compare(changed, other, opt)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "styles", "Sheet1!A1") {
		t.Fatalf("fill-bg exception must not hide a font change, got %v", got.Diffs)
	}
}

func TestCompareAnchorArraySpellingException(t *testing.T) {
	hash := formulaXLSX(t, "_xlfn.TAKE(A1#,3)", "1")
	xlfn := formulaXLSX(t, "_xlfn.TAKE(_xlfn.ANCHORARRAY(A1),3)", "1")
	opt := Options{Ignore: []string{IgnoreAnchorArraySpelling}}

	got, err := Compare(hash, xlfn, opt)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("A1# vs ANCHORARRAY must be equal with %s, got %v", IgnoreAnchorArraySpelling, got.Diffs)
	}

	got, err = Compare(hash, xlfn)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "formulas", "Sheet1!A1") {
		t.Fatalf("without exception want formulas mismatch, got %v", got.Diffs)
	}

	other := formulaXLSX(t, "_xlfn.TAKE(_xlfn.ANCHORARRAY(A1),3)", "99")
	got, err = Compare(hash, other, opt)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "values", "Sheet1!A1") {
		t.Fatalf("spelling exception must not hide a value change, got %v", got.Diffs)
	}
}

func TestCompareVolatileValuesException(t *testing.T) {
	a := formulaXLSX(t, "NOW()", "1")
	b := formulaXLSX(t, "NOW()", "2")
	opt := Options{Ignore: []string{IgnoreVolatileValues}}

	got, err := Compare(a, b, opt)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("NOW caches must be equal with %s, got %v", IgnoreVolatileValues, got.Diffs)
	}

	got, err = Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "values", "Sheet1!A1") {
		t.Fatalf("without exception want values mismatch, got %v", got.Diffs)
	}

	other := formulaXLSX(t, "TODAY()", "2")
	got, err = Compare(a, other, opt)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "formulas", "Sheet1!A1") {
		t.Fatalf("volatile-value skip must still compare formulas, got %v", got.Diffs)
	}
}

func TestCompareVolatileValuesSkipsRandArraySpill(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData>` +
			`<row r="1"><c r="A1"><f t="array" ref="A1:A2">_xlfn.RANDARRAY(2)</f><v>0.1</v></c></row>` +
			`<row r="2"><c r="A2"><v>0.2</v></c></row>` +
			`</sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData>` +
			`<row r="1"><c r="A1"><f t="array" ref="A1:A2">_xlfn.RANDARRAY(2)</f><v>0.8</v></c></row>` +
			`<row r="2"><c r="A2"><v>0.9</v></c></row>` +
			`</sheetData></worksheet>`,
	})
	opt := Options{Ignore: []string{IgnoreVolatileValues}}
	got, err := Compare(a, b, opt)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("RANDARRAY spill caches must be skipped, got %v", got.Diffs)
	}
}

func TestCompareCellFilenamePrefixException(t *testing.T) {
	a := formulaStringXLSX(t, `CELL("filename")`, `C:\old\[book.xlsx]Sheet1`)
	b := formulaStringXLSX(t, `CELL("filename")`, `/tmp/[book.xlsx]Sheet1`)
	opt := Options{Ignore: []string{IgnoreCellFilenamePrefix}}

	got, err := Compare(a, b, opt)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("CELL filename directory prefix must be equal with %s, got %v", IgnoreCellFilenamePrefix, got.Diffs)
	}

	got, err = Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "values", "Sheet1!A1") {
		t.Fatalf("without exception want values mismatch, got %v", got.Diffs)
	}

	other := formulaStringXLSX(t, `CELL("filename")`, `/tmp/[other.xlsx]Sheet1`)
	got, err = Compare(a, other, opt)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "values", "Sheet1!A1") {
		t.Fatalf("filename exception must not hide a different workbook name, got %v", got.Diffs)
	}
}

func TestCompareCellRange(t *testing.T) {
	in := numberXLSX(t, "5")
	other := numberXLSX(t, "7")
	opt := Options{Cells: map[string]CellBand{"Sheet1!A1": band(0, 10)}}

	got, err := Compare(in, other, opt)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("in-range numbers must not be a values FAIL, got %v", got.Diffs)
	}

	out := numberXLSX(t, "15")
	got, err = Compare(out, other, opt)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "values", "Sheet1!A1") {
		t.Fatalf("out of range must be a values FAIL, got %v", got.Diffs)
	}
	if !strings.Contains(got.Diffs[0].Detail, "between") {
		t.Fatalf("out-of-range detail = %q", got.Diffs[0].Detail)
	}

	eqOut := numberXLSX(t, "15")
	got, err = Compare(eqOut, numberXLSX(t, "15"), opt)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "values", "Sheet1!A1") {
		t.Fatalf("equal but outside range must still FAIL, got %v", got.Diffs)
	}

	fa := formulaXLSX(t, "A1+1", "5")
	fb := formulaXLSX(t, "A1+2", "7")
	got, err = Compare(fa, fb, opt)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal || !hasAxis(got, "formulas", "Sheet1!A1") {
		t.Fatalf("range must not hide a formula mismatch, got %v", got.Diffs)
	}
	if hasAxis(got, "values", "Sheet1!A1") {
		t.Fatalf("in-range value must not also FAIL values: %v", got.Diffs)
	}
}

func hasAxis(r Result, axis, loc string) bool {
	for _, d := range r.Diffs {
		if d.Axis == axis && d.Location == loc {
			return true
		}
	}
	return false
}

func formulaXLSX(t *testing.T, formula, value string) []byte {
	t.Helper()
	return mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1"><f>` + formula + `</f><v>` + value + `</v></c></row></sheetData></worksheet>`,
	})
}

func formulaStringXLSX(t *testing.T, formula, value string) []byte {
	t.Helper()
	return mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="str"><f>` + formula + `</f><v>` + value + `</v></c></row></sheetData></worksheet>`,
	})
}

func numberXLSX(t *testing.T, value string) []byte {
	t.Helper()
	return mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1"><v>` + value + `</v></c></row></sheetData></worksheet>`,
	})
}

func solidFillXLSX(t *testing.T, bg string) []byte {
	t.Helper()
	return mustXLSX(t, map[string]string{
		"xl/styles.xml":            solidFillStyles(bg, `<fonts count="1"><font><sz val="12"/><name val="Calibri"/></font></fonts>`),
		"xl/worksheets/sheet1.xml": solidFillSheet(),
	})
}

func solidFillSheet() string {
	return `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>y</t></is></c></row></sheetData></worksheet>`
}

func solidFillStyles(bg, fonts string) string {
	bgEl := ""
	if bg == "indexed:64" {
		bgEl = `<bgColor indexed="64"/>`
	} else if bg != "" {
		bgEl = `<bgColor rgb="` + bg + `"/>`
	}
	fills := `<fills count="2"><fill><patternFill patternType="none"/></fill>` +
		`<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/>` + bgEl + `</patternFill></fill></fills>`
	xfs := `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
		`<xf numFmtId="0" fontId="0" fillId="1" borderId="0"/></cellXfs>`
	return stylesXMLWith(fonts, fills, emptySidesBorder, xfs)
}
