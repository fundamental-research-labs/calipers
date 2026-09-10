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
