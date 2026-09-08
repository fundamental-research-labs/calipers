package compare

import (
	"archive/zip"
	"bytes"
	"os"
	"path/filepath"
	"sort"
	"testing"
	"time"
)

func TestArchivesIgnoresVolatileBits(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"[Content_Types].xml":                     contentTypes(true),
		"xl/workbook.xml":                         workbookXML("111", "0"),
		"xl/_rels/workbook.xml.rels":              workbookRels(true),
		"xl/worksheets/sheet1.xml":                sheetXML("hello"),
		"xl/calcChain.xml":                        `<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>`,
		"xl/printerSettings/printerSettings1.bin": "printer-a",
		"docProps/core.xml":                       coreXML("Alice", "2020-01-01T00:00:00Z"),
		"docProps/app.xml":                        appXML("Excel", "16.0300"),
	})
	b := mustXLSX(t, map[string]string{
		"[Content_Types].xml":        contentTypes(false),
		"xl/workbook.xml":            workbookXML("222", "0"),
		"xl/_rels/workbook.xml.rels": workbookRels(false),
		"xl/worksheets/sheet1.xml":   sheetXML("hello"),
		"docProps/core.xml":          coreXML("Bob", "2024-12-31T23:59:59Z"),
		"docProps/app.xml":           appXML("Mog", "0.1"),
	})
	got, err := Archives(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("volatile-only diffs should be equal, got %v", got.Diffs)
	}
}

func TestArchivesIgnoresPrinterSettingsOnlySheetRels(t *testing.T) {
	// Excel goldens often keep xl/worksheets/_rels/sheetN.xml.rels whose only
	// relationship is printerSettings, plus a Content_Types Default for .bin.
	// Those leftovers must not fail compare against an export that omits them.
	sheet := sheetXML("hello")
	wb := workbookXML("1", "0")
	rels := workbookRels(false)
	export := mustXLSX(t, map[string]string{
		"[Content_Types].xml":        contentTypes(false),
		"xl/workbook.xml":            wb,
		"xl/_rels/workbook.xml.rels": rels,
		"xl/worksheets/sheet1.xml":   sheet,
	})
	golden := mustXLSX(t, map[string]string{
		"[Content_Types].xml": contentTypes(false)[:len(contentTypes(false))-len("</Types>")] +
			`<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>` +
			`<Override PartName="/xl/printerSettings/printerSettings1.bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>` +
			`</Types>`,
		"xl/workbook.xml":            wb,
		"xl/_rels/workbook.xml.rels": rels,
		"xl/worksheets/sheet1.xml":   sheet,
		"xl/worksheets/_rels/sheet1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>
</Relationships>`,
		"xl/printerSettings/printerSettings1.bin": "printer-blob",
	})
	got, err := Archives(export, golden)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("printer-settings-only leftovers should be ignored, got %v", got.Diffs)
	}
	rev, err := Archives(golden, export)
	if err != nil {
		t.Fatal(err)
	}
	if !rev.Equal {
		t.Fatalf("extra printer-settings-only parts in export should be ignored, got %v", rev.Diffs)
	}
}

func TestArchivesIgnoresZipMtimes(t *testing.T) {
	parts := map[string]string{
		"[Content_Types].xml":      contentTypes(false),
		"xl/workbook.xml":          workbookXML("1", "0"),
		"xl/worksheets/sheet1.xml": sheetXML("hello"),
	}
	a := mustXLSXAt(t, parts, time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC))
	b := mustXLSXAt(t, parts, time.Date(2024, 6, 1, 12, 0, 0, 0, time.UTC))
	if bytes.Equal(a, b) {
		t.Fatal("test setup: zip bytes should differ by mtime")
	}
	got, err := Archives(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("mtime-only zips should be equal, got %v", got.Diffs)
	}
}

func TestArchivesUnequalOnCellValue(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"[Content_Types].xml":      contentTypes(false),
		"xl/workbook.xml":          workbookXML("1", "0"),
		"xl/worksheets/sheet1.xml": sheetXML("hello"),
	})
	b := mustXLSX(t, map[string]string{
		"[Content_Types].xml":      contentTypes(false),
		"xl/workbook.xml":          workbookXML("1", "0"),
		"xl/worksheets/sheet1.xml": sheetXML("world"),
	})
	got, err := Archives(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("cell value change should be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Part == "xl/worksheets/sheet1.xml" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want sheet1.xml diff, got %v", got.Diffs)
	}
}

func TestArchivesUnequalOnDate1904(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"[Content_Types].xml":      contentTypes(false),
		"xl/workbook.xml":          workbookXML("1", "0"),
		"xl/worksheets/sheet1.xml": sheetXML("hello"),
	})
	b := mustXLSX(t, map[string]string{
		"[Content_Types].xml":      contentTypes(false),
		"xl/workbook.xml":          workbookXML("1", "1"),
		"xl/worksheets/sheet1.xml": sheetXML("hello"),
	})
	got, err := Archives(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("date1904 change should be unequal")
	}
}

func TestFiles(t *testing.T) {
	dir := t.TempDir()
	parts := map[string]string{
		"[Content_Types].xml":      contentTypes(false),
		"xl/workbook.xml":          workbookXML("1", "0"),
		"xl/worksheets/sheet1.xml": sheetXML("hello"),
	}
	export := filepath.Join(dir, "export.xlsx")
	golden := filepath.Join(dir, "golden.xlsx")
	if err := os.WriteFile(export, mustXLSX(t, parts), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(golden, mustXLSX(t, parts), 0o644); err != nil {
		t.Fatal(err)
	}
	got, err := Files(export, golden)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("identical files: %v", got.Diffs)
	}
	if _, err := Files(export, filepath.Join(dir, "missing.xlsx")); err == nil {
		t.Fatal("missing golden should error")
	}
}

func mustXLSX(t *testing.T, parts map[string]string) []byte {
	t.Helper()
	return mustXLSXAt(t, parts, time.Time{})
}

func mustXLSXAt(t *testing.T, parts map[string]string, mod time.Time) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	names := make([]string, 0, len(parts))
	for name := range parts {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		h := &zip.FileHeader{Name: name, Method: zip.Deflate}
		if !mod.IsZero() {
			h.Modified = mod
		}
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

func contentTypes(withCalc bool) string {
	s := `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
	if withCalc {
		s += `<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>`
		s += `<Override PartName="/xl/printerSettings/printerSettings1.bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>`
	}
	s += `</Types>`
	return s
}

func workbookXML(calcID, date1904 string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr calcId="` + calcID + `" date1904="` + date1904 + `"/>
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
}

func workbookRels(withCalc bool) string {
	s := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>`
	if withCalc {
		s += `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>`
		s += `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="printerSettings/printerSettings1.bin"/>`
	}
	s += `</Relationships>`
	return s
}

func sheetXML(value string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>` + value + `</t></is></c></row></sheetData>
</worksheet>`
}

func coreXML(creator, created string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
<dc:creator>` + creator + `</dc:creator>
<cp:lastModifiedBy>` + creator + `</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">` + created + `</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">` + created + `</dcterms:modified>
</cp:coreProperties>`
}

func appXML(app, ver string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>` + app + `</Application>
<AppVersion>` + ver + `</AppVersion>
</Properties>`
}
