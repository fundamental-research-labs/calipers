package main

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/fundamental-research-labs/calipers/internal/cases"
)

type fakeEngine struct {
	opens   [][2]string
	scripts [][3]string
}

func (f *fakeEngine) OpenSave(in, out string) error {
	f.opens = append(f.opens, [2]string{in, out})
	return copyFile(in, out)
}

func (f *fakeEngine) RunScript(in, script, out string) error {
	f.scripts = append(f.scripts, [3]string{in, script, out})
	return copyFile(in, out)
}

func copyFile(in, out string) error {
	data, err := os.ReadFile(in)
	if err != nil {
		return err
	}
	if err := os.MkdirAll(filepath.Dir(out), 0o755); err != nil {
		return err
	}
	return os.WriteFile(out, data, 0o644)
}

func TestVerifyMissingScriptIsLoadExportOnly(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeVerifyCase(t, root, "tier_a_noscript", "hello", nil)
	fake := &fakeEngine{}
	if err := runVerify(fake, root, []string{"tier_a_noscript"}, outDir, io.Discard); err != nil {
		t.Fatal(err)
	}
	if len(fake.scripts) != 0 {
		t.Fatalf("missing script must not execute, got %v", fake.scripts)
	}
	if len(fake.opens) != 1 {
		t.Fatalf("opens=%d, want 1", len(fake.opens))
	}
	assertExportNotGolden(t, fake.opens[0][1], filepath.Join(root, "tier_a_noscript", cases.GoldenFile))
}

func TestVerifyEmptyScriptIsLoadExportOnly(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	empty := "  \n\t"
	writeVerifyCase(t, root, "tier_a_emptyjs", "hello", &empty)
	fake := &fakeEngine{}
	if err := runVerify(fake, root, []string{"tier_a_emptyjs"}, outDir, io.Discard); err != nil {
		t.Fatal(err)
	}
	if len(fake.scripts) != 0 {
		t.Fatalf("empty script must not execute, got %v", fake.scripts)
	}
	if len(fake.opens) != 1 {
		t.Fatalf("opens=%d, want 1", len(fake.opens))
	}
	assertExportNotGolden(t, fake.opens[0][1], filepath.Join(root, "tier_a_emptyjs", cases.GoldenFile))
}

func TestVerifyNonEmptyScriptRunsBeforeExport(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeVerifyCase(t, root, "tier_a_js", "hello", &js)
	fake := &fakeEngine{}
	if err := runVerify(fake, root, []string{"tier_a_js"}, outDir, io.Discard); err != nil {
		t.Fatal(err)
	}
	if len(fake.opens) != 0 {
		t.Fatalf("scripted case should not OpenSave, got %v", fake.opens)
	}
	if len(fake.scripts) != 1 {
		t.Fatalf("scripts=%d, want 1", len(fake.scripts))
	}
	if filepath.Base(fake.scripts[0][1]) != cases.ScriptFile {
		t.Fatalf("script path = %s", fake.scripts[0][1])
	}
	assertExportNotGolden(t, fake.scripts[0][2], filepath.Join(root, "tier_a_js", cases.GoldenFile))
}

func TestVerifyExportPathIsNotGolden(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeVerifyCase(t, root, "tier_a_plain", "hello", nil)
	var buf bytes.Buffer
	fake := &fakeEngine{}
	if err := runVerify(fake, root, []string{"tier_a_plain"}, outDir, &buf); err != nil {
		t.Fatal(err)
	}
	export := fake.opens[0][1]
	if filepath.Base(export) == cases.GoldenFile {
		t.Fatalf("export basename is golden: %s", export)
	}
	if filepath.Base(export) != "tier_a_plain.xlsx" {
		t.Fatalf("export = %s", export)
	}
	if !strings.Contains(buf.String(), "PASS") {
		t.Fatalf("output = %s", buf.String())
	}
}

func TestVerifyDefaultPassSkipsTierC(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeVerifyCase(t, root, "tier_a_one", "hello", nil)
	writeVerifyCase(t, root, "tier_c_bomb", "hello", nil)
	fake := &fakeEngine{}
	if err := runVerify(fake, root, nil, outDir, io.Discard); err != nil {
		t.Fatal(err)
	}
	if len(fake.opens) != 1 || !strings.Contains(fake.opens[0][0], "tier_a_one") {
		t.Fatalf("opens = %v", fake.opens)
	}
}

func TestVerifyCellValueFailsCompare(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	dir := filepath.Join(root, "tier_a_mismatch")
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, cases.InitFile), miniXLSX("hello"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, cases.GoldenFile), miniXLSX("world"), 0o644); err != nil {
		t.Fatal(err)
	}
	fake := &fakeEngine{}
	err := runVerify(fake, root, []string{"tier_a_mismatch"}, outDir, io.Discard)
	if err == nil {
		t.Fatal("value mismatch should fail the walk")
	}
}

func TestRunVerifyHelp(t *testing.T) {
	if err := run([]string{"verify", "-h"}); err != nil {
		t.Fatal(err)
	}
}

func TestRunVerifyEmptyDir(t *testing.T) {
	if err := run([]string{"verify", "--cases-dir", t.TempDir()}); err != nil {
		t.Fatal(err)
	}
}

func TestRunVerifyCLIUsesWalk(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeVerifyCase(t, root, "tier_a_cli", "hello", nil)
	fake := &fakeEngine{}
	old := newVerifyEngine
	newVerifyEngine = func() engine { return fake }
	defer func() { newVerifyEngine = old }()

	var buf bytes.Buffer
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	oldStdout := os.Stdout
	os.Stdout = w
	err = run([]string{"verify", "--cases-dir", root, "--out-dir", outDir, "--case", "tier_a_cli"})
	_ = w.Close()
	os.Stdout = oldStdout
	out, _ := io.ReadAll(r)
	_ = r.Close()
	buf.Write(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(fake.opens) != 1 {
		t.Fatalf("CLI did not walk OpenSave: %v", fake.opens)
	}
	if filepath.Base(fake.opens[0][1]) == cases.GoldenFile {
		t.Fatal("CLI wrote the golden")
	}
	if !bytes.Contains(buf.Bytes(), []byte("PASS")) {
		t.Fatalf("CLI output = %s", buf.Bytes())
	}
}

func setupVerifyDir(t *testing.T) (root, outDir string) {
	t.Helper()
	return t.TempDir(), t.TempDir()
}

func writeVerifyCase(t *testing.T, root, id, value string, script *string) {
	t.Helper()
	dir := filepath.Join(root, id)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	xlsx := miniXLSX(value)
	if err := os.WriteFile(filepath.Join(dir, cases.InitFile), xlsx, 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, cases.GoldenFile), xlsx, 0o644); err != nil {
		t.Fatal(err)
	}
	if script != nil {
		if err := os.WriteFile(filepath.Join(dir, cases.ScriptFile), []byte(*script), 0o644); err != nil {
			t.Fatal(err)
		}
	}
}

func assertExportNotGolden(t *testing.T, export, golden string) {
	t.Helper()
	if export == "" {
		t.Fatal("missing export path")
	}
	if filepath.Clean(export) == filepath.Clean(golden) {
		t.Fatalf("export path is golden: %s", export)
	}
	if filepath.Base(export) == cases.GoldenFile {
		t.Fatalf("export basename is %s", export)
	}
}

func miniXLSX(value string) []byte {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	parts := map[string]string{
		"[Content_Types].xml":      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
		"xl/workbook.xml":          `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><workbookPr calcId="1" date1904="0"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>` + value + `</t></is></c></row></sheetData></worksheet>`,
	}
	for name, body := range parts {
		w, err := zw.Create(name)
		if err != nil {
			panic(err)
		}
		if _, err := w.Write([]byte(body)); err != nil {
			panic(err)
		}
	}
	if err := zw.Close(); err != nil {
		panic(err)
	}
	return buf.Bytes()
}
