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
	writeVerifyCase(t, root, "noscript", "hello", nil)
	fake := &fakeEngine{}
	if err := runVerify(fake, root, []string{"noscript"}, outDir, io.Discard); err != nil {
		t.Fatal(err)
	}
	if len(fake.scripts) != 0 {
		t.Fatalf("missing script must not execute, got %v", fake.scripts)
	}
	if len(fake.opens) != 1 {
		t.Fatalf("opens=%d, want 1", len(fake.opens))
	}
	assertExportNotGolden(t, fake.opens[0][1], filepath.Join(root, "noscript", cases.GoldenFile))
}

func TestVerifyEmptyScriptIsLoadExportOnly(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	empty := "  \n\t"
	writeVerifyCase(t, root, "emptyjs", "hello", &empty)
	fake := &fakeEngine{}
	if err := runVerify(fake, root, []string{"emptyjs"}, outDir, io.Discard); err != nil {
		t.Fatal(err)
	}
	if len(fake.scripts) != 0 {
		t.Fatalf("empty script must not execute, got %v", fake.scripts)
	}
	if len(fake.opens) != 1 {
		t.Fatalf("opens=%d, want 1", len(fake.opens))
	}
	assertExportNotGolden(t, fake.opens[0][1], filepath.Join(root, "emptyjs", cases.GoldenFile))
}

func TestVerifyNonEmptyScriptRunsBeforeExport(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeVerifyCase(t, root, "js", "hello", &js)
	fake := &fakeEngine{}
	if err := runVerify(fake, root, []string{"js"}, outDir, io.Discard); err != nil {
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
	assertExportNotGolden(t, fake.scripts[0][2], filepath.Join(root, "js", cases.GoldenFile))
}

func TestVerifyExportPathIsNotGolden(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeVerifyCase(t, root, "plain", "hello", nil)
	var buf bytes.Buffer
	fake := &fakeEngine{}
	if err := runVerify(fake, root, []string{"plain"}, outDir, &buf); err != nil {
		t.Fatal(err)
	}
	export := fake.opens[0][1]
	if filepath.Base(export) == cases.GoldenFile {
		t.Fatalf("export basename is golden: %s", export)
	}
	if filepath.Base(export) != "plain.xlsx" {
		t.Fatalf("export = %s", export)
	}
	if !strings.Contains(buf.String(), "PASS (package match)") {
		t.Fatalf("output = %s", buf.String())
	}
}

func TestVerifyDefaultPassRunsCasesWithGoldens(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeVerifyCase(t, root, "one", "hello", nil)
	writeVerifyCase(t, root, "two", "hello", nil)
	fake := &fakeEngine{}
	if err := runVerify(fake, root, nil, outDir, io.Discard); err != nil {
		t.Fatal(err)
	}
	if len(fake.opens) != 2 {
		t.Fatalf("opens = %v, want both cases with goldens", fake.opens)
	}
}

func TestVerifyCellValueFailsCompare(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	dir := filepath.Join(root, "mismatch")
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
	var buf bytes.Buffer
	err := runVerify(fake, root, []string{"mismatch"}, outDir, &buf)
	if err == nil {
		t.Fatal("value mismatch should fail the walk")
	}
	if !strings.Contains(buf.String(), "FAIL (differing package parts: 1) xl/worksheets/sheet1.xml") {
		t.Fatalf("output must identify package-part differences: %s", buf.String())
	}
}

func TestRunVerifyHelp(t *testing.T) {
	if err := run([]string{"verify", "-h"}); err != nil {
		t.Fatal(err)
	}
}

func TestRunVerifyEmptyDir(t *testing.T) {
	if err := run([]string{"verify", "--engine", "dummy", "--cases-dir", t.TempDir()}); err != nil {
		t.Fatal(err)
	}
}

func TestHostFromSpecExcel(t *testing.T) {
	h, err := hostFromSpec("excel", false)
	if err != nil {
		t.Fatal(err)
	}
	if err := h.OpenSave("in.xlsx", "out.xlsx"); err == nil {
		t.Fatal("excel host off Windows should error on OpenSave")
	}
}

func TestVerifyRecalculateRejectsExcelBeforeCreatingHost(t *testing.T) {
	old := newExcelHost
	newExcelHost = func() engine {
		t.Fatal("unsupported recalculation must fail before starting Excel")
		return nil
	}
	t.Cleanup(func() { newExcelHost = old })
	err := run([]string{"verify", "--engine", "excel", "--recalculate", "--cases-dir", t.TempDir()})
	if err == nil || !strings.Contains(err.Error(), "--recalculate is unsupported with --engine excel") {
		t.Fatalf("error = %v, want unsupported Excel calculation policy", err)
	}
}

func TestVerifyCLIRecalculatePolicy(t *testing.T) {
	for _, flag := range []string{"--recalculate", "--recalculate=false"} {
		t.Run(flag, func(t *testing.T) {
			want := flag == "--recalculate"
			old := newBinaryHost
			called := false
			newBinaryHost = func(path string, recalculate bool) engine {
				called = true
				if path != "engine-under-test" || recalculate != want {
					t.Fatalf("host options = %q, %v; want engine-under-test, %v", path, recalculate, want)
				}
				return &fakeEngine{}
			}
			t.Cleanup(func() { newBinaryHost = old })
			r, w, err := os.Pipe()
			if err != nil {
				t.Fatal(err)
			}
			oldStdout := os.Stdout
			os.Stdout = w
			err = run([]string{"verify", "--engine", "engine-under-test", flag, "--cases-dir", t.TempDir(), "--out-dir", t.TempDir()})
			os.Stdout = oldStdout
			_ = w.Close()
			out, readErr := io.ReadAll(r)
			_ = r.Close()
			if err != nil || readErr != nil || !called {
				t.Fatalf("CLI error = %v, read error = %v, host constructed = %v", err, readErr, called)
			}
			policy := "host default (no recalculation requested)"
			if want {
				policy = "recalculate before export (external engine --recalculate)"
			}
			if !strings.Contains(string(out), "calculation policy: "+policy) {
				t.Fatalf("missing selected policy in output: %s", out)
			}
		})
	}
}

func TestVerifyRequiresEngineEvenWhenMogBinSet(t *testing.T) {
	t.Setenv("MOG_BIN", "/pretend/mog")
	t.Setenv("ENGINE", "/pretend/engine")
	err := run([]string{"verify", "--cases-dir", t.TempDir()})
	if err == nil || !strings.Contains(err.Error(), "verify requires --engine") {
		t.Fatalf("error = %v, want required --engine", err)
	}
	if strings.Contains(err.Error(), "MOG_BIN") || strings.Contains(err.Error(), "vendor/mog") {
		t.Fatalf("error must not suggest env/vendor fallback: %v", err)
	}
}

func TestHostFromSpecRequiresEngine(t *testing.T) {
	t.Setenv("MOG_BIN", "/pretend/mog")
	t.Setenv("ENGINE", "/pretend/engine")
	_, err := hostFromSpec("", false)
	if err == nil || !strings.Contains(err.Error(), "verify requires --engine") {
		t.Fatalf("error = %v, want required --engine", err)
	}
	if strings.Contains(err.Error(), "MOG_BIN") || strings.Contains(err.Error(), "vendor/mog") {
		t.Fatalf("error must not suggest env/vendor fallback: %v", err)
	}
}

func TestHostFromSpecPathConstructsBinaryHost(t *testing.T) {
	t.Setenv("MOG_BIN", "/from-env/mog")
	old := newBinaryHost
	var gotPath string
	var gotRecalc bool
	newBinaryHost = func(path string, recalculate bool) engine {
		gotPath = path
		gotRecalc = recalculate
		return &fakeEngine{}
	}
	t.Cleanup(func() { newBinaryHost = old })
	h, err := hostFromSpec("/path/to/engine", true)
	if err != nil {
		t.Fatal(err)
	}
	if h == nil {
		t.Fatal("nil host")
	}
	if gotPath != "/path/to/engine" || !gotRecalc {
		t.Fatalf("path=%q recalc=%v, want /path/to/engine true", gotPath, gotRecalc)
	}
}

func TestVerifyCLIEngineFlagIgnoresMogBin(t *testing.T) {
	t.Setenv("MOG_BIN", "/from-env/mog")
	old := newBinaryHost
	var gotPath string
	newBinaryHost = func(path string, recalculate bool) engine {
		gotPath = path
		return &fakeEngine{}
	}
	t.Cleanup(func() { newBinaryHost = old })
	_, err := captureStdout(t, func() error {
		return run([]string{"verify", "--engine", "/from-flag/engine", "--cases-dir", t.TempDir()})
	})
	if err != nil {
		t.Fatal(err)
	}
	if gotPath != "/from-flag/engine" {
		t.Fatalf("path = %q, want /from-flag/engine (MOG_BIN must not select the engine)", gotPath)
	}
}

func TestRunVerifyCLIUsesWalk(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeVerifyCase(t, root, "cli", "hello", nil)
	fake := &fakeEngine{}
	old := newBinaryHost
	newBinaryHost = func(path string, recalculate bool) engine {
		if recalculate {
			t.Fatal("default CLI must not request recalculation")
		}
		return fake
	}
	defer func() { newBinaryHost = old }()

	var buf bytes.Buffer
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	oldStdout := os.Stdout
	os.Stdout = w
	err = run([]string{"verify", "--engine", "dummy", "--cases-dir", root, "--out-dir", outDir, "--case", "cli"})
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

func TestRunVerifyCLIAllSuites(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeNestedVerifyCase(t, root, "roundtrip", "rt", "hello", nil)
	writeNestedVerifyCase(t, root, "default", "other", "hello", nil)
	fake := &fakeEngine{}
	restore := stubBinaryHost(t, fake)

	out, err := captureStdout(t, func() error {
		return run([]string{"verify", "--engine", "dummy", "--cases-dir", root, "--out-dir", outDir})
	})
	restore()
	if err != nil {
		t.Fatal(err)
	}
	if len(fake.opens) != 2 {
		t.Fatalf("default verify must run every suite: opens=%v", fake.opens)
	}
	got := openedCaseIDs(fake)
	if !got["rt"] || !got["other"] {
		t.Fatalf("opened = %v, want roundtrip+default", got)
	}
	if !strings.Contains(out, "roundtrip/rt") || !strings.Contains(out, "default/other") {
		t.Fatalf("output must name suite-qualified ids:\n%s", out)
	}
}

func TestRunVerifyCLICasesDirSingleSuite(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeNestedVerifyCase(t, root, "roundtrip", "rt", "hello", nil)
	writeNestedVerifyCase(t, root, "default", "other", "hello", nil)
	fake := &fakeEngine{}
	restore := stubBinaryHost(t, fake)

	_, err := captureStdout(t, func() error {
		return run([]string{"verify", "--engine", "dummy", "--cases-dir", filepath.Join(root, "roundtrip"), "--out-dir", outDir})
	})
	restore()
	if err != nil {
		t.Fatal(err)
	}
	got := openedCaseIDs(fake)
	if len(fake.opens) != 1 || !got["rt"] || got["other"] {
		t.Fatalf("--cases-dir <suite> opened %v", got)
	}
}

func TestRunVerifyCLISuiteFlag(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeNestedVerifyCase(t, root, "roundtrip", "rt", "hello", nil)
	writeNestedVerifyCase(t, root, "default", "other", "hello", nil)
	fake := &fakeEngine{}
	restore := stubBinaryHost(t, fake)

	out, err := captureStdout(t, func() error {
		return run([]string{"verify", "--engine", "dummy", "--cases-dir", root, "--out-dir", outDir, "--suite", "roundtrip"})
	})
	restore()
	if err != nil {
		t.Fatal(err)
	}
	got := openedCaseIDs(fake)
	if len(fake.opens) != 1 || !got["rt"] || got["other"] {
		t.Fatalf("--suite roundtrip opened %v", got)
	}
	if !strings.Contains(out, "roundtrip/rt") {
		t.Fatalf("output = %s", out)
	}
	if strings.Contains(out, "other") {
		t.Fatalf("--suite must not run other suites:\n%s", out)
	}
}

func TestRunVerifyCLICaseFlag(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeNestedVerifyCase(t, root, "roundtrip", "rt", "hello", nil)
	writeNestedVerifyCase(t, root, "default", "other", "hello", nil)
	fake := &fakeEngine{}
	restore := stubBinaryHost(t, fake)

	out, err := captureStdout(t, func() error {
		return run([]string{"verify", "--engine", "dummy", "--cases-dir", root, "--out-dir", outDir, "--case", "default/other"})
	})
	restore()
	if err != nil {
		t.Fatal(err)
	}
	got := openedCaseIDs(fake)
	if len(fake.opens) != 1 || !got["other"] || got["rt"] {
		t.Fatalf("--case opened %v", got)
	}
	if !strings.Contains(out, "default/other") {
		t.Fatalf("output = %s", out)
	}
	if filepath.Base(filepath.Dir(fake.opens[0][1])) != "default" {
		t.Fatalf("export must be under suite dir, got %s", fake.opens[0][1])
	}
}

func TestRunVerifyCLISuiteAndCase(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeNestedVerifyCase(t, root, "roundtrip", "rt", "hello", nil)
	writeNestedVerifyCase(t, root, "default", "other", "hello", nil)
	fake := &fakeEngine{}
	restore := stubBinaryHost(t, fake)

	_, err := captureStdout(t, func() error {
		return run([]string{"verify", "--engine", "dummy", "--cases-dir", root, "--out-dir", outDir, "--suite", "roundtrip", "--case", "roundtrip/rt"})
	})
	restore()
	if err != nil {
		t.Fatal(err)
	}
	got := openedCaseIDs(fake)
	if len(fake.opens) != 1 || !got["rt"] {
		t.Fatalf("--suite --case opened %v", got)
	}

	fake2 := &fakeEngine{}
	restore = stubBinaryHost(t, fake2)
	err = run([]string{"verify", "--engine", "dummy", "--cases-dir", root, "--out-dir", outDir, "--suite", "roundtrip", "--case", "other"})
	restore()
	if err == nil || !strings.Contains(err.Error(), "unknown case") {
		t.Fatalf("error = %v, want unknown case outside suite", err)
	}
}

func TestRunVerifyCLIDefaultSkipsCasesWithoutGolden(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeNestedVerifyCase(t, root, "roundtrip", "rt", "hello", nil)
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeNestedVerifyCaseNoGolden(t, root, "scratch", "text", "hello", &js)
	fake := &fakeEngine{}
	restore := stubBinaryHost(t, fake)

	out, err := captureStdout(t, func() error {
		return run([]string{"verify", "--engine", "dummy", "--cases-dir", root, "--out-dir", outDir})
	})
	restore()
	if err != nil {
		t.Fatal(err)
	}
	if len(fake.opens) != 1 || len(fake.scripts) != 0 {
		t.Fatalf("default verify must skip cases without a golden: opens=%v scripts=%v", fake.opens, fake.scripts)
	}
	got := openedCaseIDs(fake)
	if !got["rt"] || got["text"] {
		t.Fatalf("opened = %v, want only the case with a golden", got)
	}
	if strings.Contains(out, "scratch/") {
		t.Fatalf("default walk must not mention cases without a golden:\n%s", out)
	}
}

func TestRunVerifyCLISuiteScratch(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeNestedVerifyCase(t, root, "roundtrip", "rt", "hello", nil)
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeNestedVerifyCase(t, root, "scratch", "text", "hello", &js)
	fake := &fakeEngine{}
	restore := stubBinaryHost(t, fake)

	out, err := captureStdout(t, func() error {
		return run([]string{"verify", "--engine", "dummy", "--cases-dir", root, "--out-dir", outDir, "--suite", "scratch"})
	})
	restore()
	if err != nil {
		t.Fatal(err)
	}
	if len(fake.opens) != 0 || len(fake.scripts) != 1 {
		t.Fatalf("--suite scratch must run the Office.js case: opens=%v scripts=%v", fake.opens, fake.scripts)
	}
	if !strings.Contains(out, "scratch/text") {
		t.Fatalf("output must name scratch id:\n%s", out)
	}
	if strings.Contains(out, "rt") {
		t.Fatalf("--suite scratch must not run other suites:\n%s", out)
	}
}

func TestRunVerifyCLIUnknownSuite(t *testing.T) {
	root, outDir := setupVerifyDir(t)
	writeNestedVerifyCase(t, root, "roundtrip", "rt", "hello", nil)
	restore := stubBinaryHost(t, &fakeEngine{})
	err := run([]string{"verify", "--engine", "dummy", "--cases-dir", root, "--out-dir", outDir, "--suite", "nope"})
	restore()
	if err == nil || !strings.Contains(err.Error(), "unknown suite") {
		t.Fatalf("error = %v, want unknown suite", err)
	}
}

func TestRunVerifyHelpListsSuiteAndCase(t *testing.T) {
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	old := os.Stdout
	os.Stdout = w
	err = run([]string{"verify", "-h"})
	_ = w.Close()
	os.Stdout = old
	out, _ := io.ReadAll(r)
	_ = r.Close()
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(out, []byte("--suite")) || !bytes.Contains(out, []byte("--case")) {
		t.Fatalf("verify help must list --suite and --case:\n%s", out)
	}
	if !bytes.Contains(out, []byte("scratch")) {
		t.Fatalf("verify help must mention scratch:\n%s", out)
	}
	if !bytes.Contains(out, []byte("--engine is required")) {
		t.Fatalf("verify help must say --engine is required:\n%s", out)
	}
	if bytes.Contains(out, []byte("may be omitted")) {
		t.Fatalf("verify help must not say --engine may be omitted:\n%s", out)
	}
	if bytes.Contains(out, []byte("MOG_BIN")) || bytes.Contains(out, []byte("vendor/mog")) {
		t.Fatalf("verify help must not name Mog env/vendor as a default engine:\n%s", out)
	}
}

func stubBinaryHost(t *testing.T, fake engine) (restore func()) {
	t.Helper()
	old := newBinaryHost
	newBinaryHost = func(path string, recalculate bool) engine {
		if recalculate {
			t.Fatal("default CLI must not request recalculation")
		}
		return fake
	}
	return func() { newBinaryHost = old }
}

func captureStdout(t *testing.T, fn func() error) (string, error) {
	t.Helper()
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	old := os.Stdout
	os.Stdout = w
	fnErr := fn()
	_ = w.Close()
	os.Stdout = old
	out, _ := io.ReadAll(r)
	_ = r.Close()
	return string(out), fnErr
}

func openedCaseIDs(fake *fakeEngine) map[string]bool {
	got := map[string]bool{}
	for _, open := range fake.opens {
		got[filepath.Base(filepath.Dir(open[0]))] = true
	}
	return got
}

func writeNestedVerifyCase(t *testing.T, root, suite, id, value string, script *string) {
	t.Helper()
	writeVerifyCase(t, filepath.Join(root, suite), id, value, script)
}

func writeNestedVerifyCaseNoGolden(t *testing.T, root, suite, id, value string, script *string) {
	t.Helper()
	writeVerifyCaseFiles(t, filepath.Join(root, suite), id, value, script, false)
}

func setupVerifyDir(t *testing.T) (root, outDir string) {
	t.Helper()
	return t.TempDir(), t.TempDir()
}

func writeVerifyCase(t *testing.T, root, id, value string, script *string) {
	t.Helper()
	writeVerifyCaseFiles(t, root, id, value, script, true)
}

func writeVerifyCaseFiles(t *testing.T, root, id, value string, script *string, golden bool) {
	t.Helper()
	dir := filepath.Join(root, id)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	xlsx := miniXLSX(value)
	if err := os.WriteFile(filepath.Join(dir, cases.InitFile), xlsx, 0o644); err != nil {
		t.Fatal(err)
	}
	if golden {
		if err := os.WriteFile(filepath.Join(dir, cases.GoldenFile), xlsx, 0o644); err != nil {
			t.Fatal(err)
		}
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
