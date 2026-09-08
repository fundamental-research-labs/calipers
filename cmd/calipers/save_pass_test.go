package main

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
	"github.com/fundamental-research-labs/calipers/internal/golden"
)

func TestRunExcelSavePassUsage(t *testing.T) {
	err := run([]string{"excel-save-pass", "a", "b"})
	if err == nil || !strings.Contains(err.Error(), "usage:") {
		t.Fatalf("run(excel-save-pass extra args) = %v, want usage error", err)
	}
}

func TestRunExcelSavePassEmptyDir(t *testing.T) {
	dir := t.TempDir()
	if err := run([]string{"excel-save-pass", dir}); err != nil {
		t.Fatalf("empty corpus should not require Excel: %v", err)
	}
}

func TestSkippedOpenSave(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async () => {});"
	writeCaseDir(t, root, "tier_a_plain", nil)
	writeCaseDir(t, root, "tier_a_js", &js)
	writeCaseDir(t, root, "tier_c_bomb", nil)
	all, err := cases.Load(root)
	if err != nil {
		t.Fatal(err)
	}
	skipped := skippedOpenSave(all)
	if len(skipped) != 1 || skipped[0].ID != "tier_a_js" {
		t.Fatalf("skipped = %+v", skipped)
	}
	if n := len(cases.OpenSavePass(all)); n != 1 {
		t.Fatalf("open+save pass len=%d, want 1", n)
	}
}

func TestCheckUniformOpenSaveGoldens(t *testing.T) {
	root := t.TempDir()
	writeCaseDir(t, root, "tier_a_one", nil)
	writeCaseDir(t, root, "tier_a_two", nil)
	all, err := cases.Load(root)
	if err != nil {
		t.Fatal(err)
	}

	writeMeta(t, all[0].GoldenPath, golden.New(excel.HostID, "windows", "16.0", "20326", "init.xlsx"))
	writeMeta(t, all[1].GoldenPath, golden.New(excel.HostID, "windows", "16.0", "20326", "init.xlsx"))
	if err := checkUniformOpenSaveGoldens(all, nil); err != nil {
		t.Fatalf("uniform set: %v", err)
	}

	mac := golden.New("excel-mac", "darwin", "16.0", "20326", "init.xlsx")
	writeMeta(t, all[1].GoldenPath, mac)
	err = checkUniformOpenSaveGoldens(all, nil)
	if err == nil || !strings.Contains(err.Error(), "excel-mac") {
		t.Fatalf("want refuse excel-mac, got %v", err)
	}

	writeMeta(t, all[1].GoldenPath, golden.New(excel.HostID, "windows", "16.0", "99999", "init.xlsx"))
	err = checkUniformOpenSaveGoldens(all, nil)
	if err == nil || !strings.Contains(err.Error(), "mixed excel goldens") {
		t.Fatalf("want mixed goldens error, got %v", err)
	}

	// Failed cases are ignored for the uniform check.
	if err := checkUniformOpenSaveGoldens(all, map[string]error{all[1].ID: os.ErrInvalid}); err != nil {
		t.Fatalf("failed sibling should be skipped: %v", err)
	}

	scripted := golden.New(excel.HostID, "windows", "16.0", "20326", "init.xlsx")
	scripted.Script = "script.js"
	writeMeta(t, all[1].GoldenPath, scripted)
	err = checkUniformOpenSaveGoldens(all, nil)
	if err == nil || !strings.Contains(err.Error(), "omit script") {
		t.Fatalf("want omit-script error, got %v", err)
	}
}

func writeCaseDir(t *testing.T, root, id string, script *string) {
	t.Helper()
	dir := filepath.Join(root, id)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, cases.InitFile), []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	if script != nil {
		if err := os.WriteFile(filepath.Join(dir, cases.ScriptFile), []byte(*script), 0o644); err != nil {
			t.Fatal(err)
		}
	}
}

func writeMeta(t *testing.T, xlsxPath string, m golden.Meta) {
	t.Helper()
	if err := os.WriteFile(xlsxPath, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := golden.Write(xlsxPath, m); err != nil {
		t.Fatal(err)
	}
}
