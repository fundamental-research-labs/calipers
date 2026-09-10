package xlsxmodel

import (
	"archive/zip"
	"bytes"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"testing"
	"time"
)

func TestCompareEqualOnBinary64FloatText(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="4"><c r="C4"><v>92.3</v></c></row></sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="4"><c r="C4"><v>92.299999999999997</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("92.3 vs 92.299999999999997 must be equal (binary64), got %v", got.Diffs)
	}
	t.Log("equal: sheet cell 92.3 vs 92.299999999999997")
}

func TestCompareUnequalOnCellString(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("0"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	b := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("0"),
		"xl/worksheets/sheet1.xml": sheetInline("world"),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("hello vs world must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "values" && d.Location == "Sheet1!A1" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want values Sheet1!A1, got %v", got.Diffs)
	}
	t.Logf("unequal: hello vs world %v", got.Diffs)
}

func TestCompareUnequalOnDate1904(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("0"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	b := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("1"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("date1904 0 vs 1 must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "date1904" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want date1904 diff, got %v", got.Diffs)
	}
	t.Logf("unequal: date1904 0 vs 1 %v", got.Diffs)
}

func TestCompareUnequalOnDate1904NotClobberedByX15WorkbookPr(t *testing.T) {
	// Corpus Excel files emit a later x15:workbookPr with no date1904.
	// Matching Local=="workbookPr" would overwrite date1904=1 with false.
	with1904 := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML1904WithX15("1"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	without := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("0"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	got, err := Compare(with1904, without)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("date1904=1 plus x15:workbookPr vs date1904=0 must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "date1904" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want date1904 diff, got %v", got.Diffs)
	}
	t.Logf("unequal: date1904=1 with trailing x15:workbookPr vs 0 %v", got.Diffs)
}

func TestCompareEqualWhenDate1904AbsentMeans1900(t *testing.T) {
	absent := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          `<workbook><workbookPr/><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>`,
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	zero := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("0"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	got, err := Compare(absent, zero)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("absent date1904 must match 0, got %v", got.Diffs)
	}
}

func TestCompareEqualOnEmptyVsSelfClosingCompany(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"docProps/app.xml":         `<?xml version="1.0"?><Properties><Application>Mog</Application><Company/><AppVersion>0.1</AppVersion></Properties>`,
		"xl/worksheets/sheet1.xml": sheetInline("x"),
	})
	b := mustXLSX(t, map[string]string{
		"docProps/app.xml":         `<?xml version="1.0"?><Properties><Application>Excel</Application><Company></Company><AppVersion>16.0300</AppVersion></Properties>`,
		"xl/worksheets/sheet1.xml": sheetInline("x"),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("Company serialization must not fail value compare, got %v", got.Diffs)
	}
}

func TestCompareUnequalOnBooleanVsNumber(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="3"><c r="B3" t="b"><v>1</v></c></row></sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="3"><c r="B3"><v>1</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("boolean t=b vs number 1 must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "types" && d.Location == "Sheet1!B3" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want types Sheet1!B3, got %v", got.Diffs)
	}
	t.Logf("unequal: boolean vs number %v", got.Diffs)
}

func TestCompareUnequalOnFormulaText(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="C1"><f>A1+B1</f><v>15</v></c></row></sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="C1"><f>A1-B1</f><v>15</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("A1+B1 vs A1-B1 must be unequal even with the same cached v")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "formulas" && d.Location == "Sheet1!C1" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want formulas Sheet1!C1, got %v", got.Diffs)
	}
	for _, d := range got.Diffs {
		if d.Axis == "values" {
			t.Fatalf("cached value matches; extra values diff: %v", got.Diffs)
		}
	}
	t.Logf("unequal: formula text %v", got.Diffs)
}

func TestCompareUnequalOnCachedFormulaValue(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="C1"><f>A1+B1</f><v>0</v></c></row></sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="C1"><f>A1+B1</f><v>15</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("cached v=0 vs v=15 must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "values" && d.Location == "Sheet1!C1" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want values Sheet1!C1, got %v", got.Diffs)
	}
	for _, d := range got.Diffs {
		if d.Axis == "formulas" {
			t.Fatalf("formula text matches; extra formulas diff: %v", got.Diffs)
		}
	}
	t.Logf("unequal: cached formula value %v", got.Diffs)
}

func TestCompareEqualOnSharedFormulaExpansion(t *testing.T) {
	shared := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData>` +
			`<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B2</f><v>1</v></c></row>` +
			`<row r="3"><c r="A3"><f t="shared" si="0"/><v>2</v></c></row>` +
			`</sheetData></worksheet>`,
	})
	expanded := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData>` +
			`<row r="2"><c r="A2"><f>B2</f><v>1</v></c></row>` +
			`<row r="3"><c r="A3"><f>B3</f><v>2</v></c></row>` +
			`</sheetData></worksheet>`,
	})
	got, err := Compare(shared, expanded)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("shared B2 at A2/A3 must match expanded B2/B3, got %v", got.Diffs)
	}
	t.Log("equal: shared formula expansion B2 -> B3")
}

func TestCompareEqualOnSSTVsInlineStr(t *testing.T) {
	sst := mustXLSX(t, map[string]string{
		"xl/sharedStrings.xml":     `<sst><si><t>hello</t></si></sst>`,
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`,
	})
	inline := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	got, err := Compare(sst, inline)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("SST and inlineStr are the same string type, got %v", got.Diffs)
	}
}

func TestCompareFilesRoundtripFormulasInitVsGolden(t *testing.T) {
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("runtime.Caller failed")
	}
	dir := filepath.Join(filepath.Dir(file), "..", "..", "verification", "cases", "roundtrip", "formulas")
	got, err := CompareFiles(filepath.Join(dir, "init.xlsx"), filepath.Join(dir, "golden.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("formulas init v=0 vs golden cached values must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "values" && d.Location == "Sheet1!C1" {
			found = true
		}
		if d.Axis == "formulas" {
			t.Fatalf("formula text should match; got %v", got.Diffs)
		}
	}
	if !found {
		t.Fatalf("want values Sheet1!C1 (0 vs 15), got %v", got.Diffs)
	}
	t.Logf("unequal: roundtrip/formulas cached values %v", got.Diffs)
}

func TestCompareFilesRoundtripSimpleInitVsGolden(t *testing.T) {
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("runtime.Caller failed")
	}
	dir := filepath.Join(filepath.Dir(file), "..", "..", "verification", "cases", "roundtrip", "simple")
	got, err := CompareFiles(filepath.Join(dir, "init.xlsx"), filepath.Join(dir, "golden.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("roundtrip/simple init vs golden must match on values/date1904 (C4 92.3), got %v", got.Diffs)
	}
	t.Log("equal: roundtrip/simple init.xlsx vs golden.xlsx (resolved values including C4)")
}

func TestCompareFilesMissing(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "a.xlsx")
	if err := os.WriteFile(p, mustXLSX(t, map[string]string{"xl/worksheets/sheet1.xml": sheetInline("x")}), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := CompareFiles(p, filepath.Join(dir, "missing.xlsx")); err == nil {
		t.Fatal("missing golden should error")
	}
}

func workbookXML(date1904 string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr calcId="1" date1904="` + date1904 + `"/>
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
}

func workbookXML1904WithX15(date1904 string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
<workbookPr calcId="1" date1904="` + date1904 + `"/>
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
<extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}"><x15:workbookPr chartTrackingRefBase="1"/></ext></extLst>
</workbook>`
}

func sheetInline(value string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>` + value + `</t></is></c></row></sheetData>
</worksheet>`
}

func mustXLSX(t *testing.T, parts map[string]string) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	names := make([]string, 0, len(parts))
	for name := range parts {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		h := &zip.FileHeader{Name: name, Method: zip.Deflate, Modified: time.Time{}}
		w, err := zw.CreateHeader(h)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write([]byte(parts[name])); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}
