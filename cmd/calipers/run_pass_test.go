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

func TestRunExcelRunPassUsage(t *testing.T) {
	err := run([]string{"excel-run-pass", "a", "b"})
	if err == nil || !strings.Contains(err.Error(), "usage:") {
		t.Fatalf("run(excel-run-pass extra args) = %v, want usage error", err)
	}
}

func TestRunExcelRunPassEmptyDir(t *testing.T) {
	dir := t.TempDir()
	if err := run([]string{"excel-run-pass", dir}); err != nil {
		t.Fatalf("empty corpus should not require Excel: %v", err)
	}
}

func TestRunExcelRunPassSkipsExistingGoldensWithoutExcel(t *testing.T) {
	root := t.TempDir()
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeCaseDir(t, root, "js", &js)
	if err := os.WriteFile(filepath.Join(root, "js", cases.GoldenFile), []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := run([]string{"excel-run-pass", root}); err != nil {
		t.Fatalf("no pending goldens should not require Excel: %v", err)
	}
}

func TestCheckUniformScriptedGoldens(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async () => {});"
	writeCaseDir(t, root, "one", &js)
	writeCaseDir(t, root, "two", &js)
	all, err := cases.Load(root)
	if err != nil {
		t.Fatal(err)
	}

	m1 := golden.New(excel.HostID, "windows", "16.0", "20326", "init.xlsx")
	m1.Script = cases.ScriptFile
	m2 := golden.New(excel.HostID, "windows", "16.0", "20326", "init.xlsx")
	m2.Script = cases.ScriptFile
	writeMeta(t, all[0].GoldenPath, m1)
	writeMeta(t, all[1].GoldenPath, m2)
	if err := checkUniformScriptedGoldens(all, nil); err != nil {
		t.Fatalf("uniform set: %v", err)
	}

	load := golden.New(excel.HostID, "windows", "16.0", "20326", "init.xlsx")
	writeMeta(t, all[1].GoldenPath, load)
	err = checkUniformScriptedGoldens(all, nil)
	if err == nil || !strings.Contains(err.Error(), "scripted golden must record") {
		t.Fatalf("want script identity error, got %v", err)
	}
}
