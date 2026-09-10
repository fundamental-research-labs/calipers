package excel

import (
	"archive/zip"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

func TestCheckWindowsExcel16ExportAcceptsWindowsExcel16(t *testing.T) {
	path := writeAppXLSX(t, namespacedAppXML("Microsoft Excel", "16.0300"))
	if err := CheckWindowsExcel16Export(path); err != nil {
		t.Fatal(err)
	}
	app, ver, err := ReadAppProperties(path)
	if err != nil {
		t.Fatal(err)
	}
	if app != WindowsExcelApplication || !strings.HasPrefix(ver, "16.") {
		t.Fatalf("ReadAppProperties = %q %q", app, ver)
	}
}

func TestCheckWindowsExcel16ExportAcceptsPrefixedAppXML(t *testing.T) {
	xml := `<?xml version="1.0"?><ep:Properties xmlns:ep="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><ep:Application>Microsoft Excel</ep:Application><ep:AppVersion>16.0300</ep:AppVersion></ep:Properties>`
	path := writeAppXLSX(t, xml)
	if err := CheckWindowsExcel16Export(path); err != nil {
		t.Fatal(err)
	}
}

func TestCheckWindowsExcel16ExportRejectsProducers(t *testing.T) {
	cases := []struct {
		name string
		app  string
		ver  string
		want string
	}{
		{"macintosh-16", "Microsoft Macintosh Excel", "16.0300", "Microsoft Macintosh Excel"},
		{"macintosh-14", "Microsoft Macintosh Excel", "14.0300", "Microsoft Macintosh Excel"},
		{"macintosh-12", "Microsoft Macintosh Excel", "12.0000", "Microsoft Macintosh Excel"},
		{"excel-online", "Microsoft Excel Online", "16.0300", "Microsoft Excel Online"},
		{"openpyxl", "Microsoft Excel Compatible / Openpyxl 3.1.5", "3.1", "Openpyxl"},
		{"libreoffice", "LibreOffice Calc", "7.6", "LibreOffice"},
		{"node-xlsx-stream", "node-xlsx-stream", "0.1.2", "node-xlsx-stream"},
		{"odf-converter", "ODF Converter", "12.0000", "ODF Converter"},
		{"poi-12", "Microsoft Excel", "12.0000", "12.0000"},
		{"excel-14", "Microsoft Excel", "14.0300", "14.0300"},
		{"excel-15", "Microsoft Excel", "15.0300", "15.0300"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			path := writeAppXLSX(t, namespacedAppXML(tc.app, tc.ver))
			err := CheckWindowsExcel16Export(path)
			if err == nil {
				t.Fatalf("accepted Application=%q AppVersion=%q", tc.app, tc.ver)
			}
			if !strings.Contains(err.Error(), tc.want) {
				t.Fatalf("error %q, want substring %q", err, tc.want)
			}
		})
	}
}

func TestCheckWindowsExcel16ExportRejectsMissingApplication(t *testing.T) {
	xml := `<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><AppVersion>16.0300</AppVersion></Properties>`
	path := writeAppXLSX(t, xml)
	err := CheckWindowsExcel16Export(path)
	if err == nil || !strings.Contains(err.Error(), "missing Application") {
		t.Fatalf("error = %v, want missing Application", err)
	}
}

func TestCheckWindowsExcel16ExportRejectsMissingAppVersion(t *testing.T) {
	xml := `<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>`
	path := writeAppXLSX(t, xml)
	err := CheckWindowsExcel16Export(path)
	if err == nil || !strings.Contains(err.Error(), "missing AppVersion") {
		t.Fatalf("error = %v, want missing AppVersion", err)
	}
}

func TestCheckWindowsExcel16ExportRejectsMissingAppXML(t *testing.T) {
	path := filepath.Join(t.TempDir(), "init.xlsx")
	writeZipFile(t, path, map[string]string{"[Content_Types].xml": "<Types/>"})
	err := CheckWindowsExcel16Export(path)
	if err == nil || !strings.Contains(err.Error(), "missing docProps/app.xml") {
		t.Fatalf("error = %v, want missing docProps/app.xml", err)
	}
}

func TestCheckWindowsExcel16ExportRejectsNonZip(t *testing.T) {
	path := filepath.Join(t.TempDir(), "init.xlsx")
	if err := os.WriteFile(path, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	err := CheckWindowsExcel16Export(path)
	if err == nil || !strings.Contains(err.Error(), "not a zip workbook") {
		t.Fatalf("error = %v, want not a zip workbook", err)
	}
}

func namespacedAppXML(app, ver string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
		`<Application>` + app + `</Application>` +
		`<AppVersion>` + ver + `</AppVersion>` +
		`</Properties>`
}

func writeAppXLSX(t *testing.T, appXML string) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "init.xlsx")
	writeZipFile(t, path, map[string]string{"docProps/app.xml": appXML})
	return path
}

func writeZipFile(t *testing.T, path string, parts map[string]string) {
	t.Helper()
	f, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	zw := zip.NewWriter(f)
	for name, body := range parts {
		h := &zip.FileHeader{Name: name, Method: zip.Deflate, Modified: time.Time{}}
		w, err := zw.CreateHeader(h)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write([]byte(body)); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}
}
