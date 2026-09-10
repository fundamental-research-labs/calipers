package main

import (
	"bytes"
	"errors"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"

	"github.com/fundamental-research-labs/calipers/internal/excel"
	"github.com/fundamental-research-labs/calipers/internal/golden"
)

func TestRunHelp(t *testing.T) {
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	old := os.Stdout
	os.Stdout = w
	err = run(nil)
	_ = w.Close()
	os.Stdout = old
	out, _ := io.ReadAll(r)
	_ = r.Close()
	if err != nil {
		t.Fatalf("run(nil) = %v", err)
	}
	if !bytes.Contains(out, []byte("excel-run")) || !bytes.Contains(out, []byte("excel-save")) {
		t.Fatalf("help must list excel-run and excel-save:\n%s", out)
	}
	if !bytes.Contains(out, []byte("excel-save-pass")) {
		t.Fatalf("help must list excel-save-pass:\n%s", out)
	}
	if !bytes.Contains(out, []byte("excel-run-pass")) {
		t.Fatalf("help must list excel-run-pass:\n%s", out)
	}
	if !bytes.Contains(out, []byte("measure-budgets")) {
		t.Fatalf("help must list measure-budgets:\n%s", out)
	}
	if !bytes.Contains(out, []byte("verify")) {
		t.Fatalf("help must list verify:\n%s", out)
	}
	if !bytes.Contains(out, []byte("--suite")) || !bytes.Contains(out, []byte("--case")) {
		t.Fatalf("help must list --suite and --case:\n%s", out)
	}
	if !bytes.Contains(out, []byte("scratch")) {
		t.Fatalf("help must mention scratch:\n%s", out)
	}
	if !bytes.Contains(out, []byte("--engine is required")) {
		t.Fatalf("help must say --engine is required:\n%s", out)
	}
	if bytes.Contains(out, []byte("may be omitted")) {
		t.Fatalf("help must not say --engine may be omitted:\n%s", out)
	}
	if bytes.Contains(out, []byte("MOG_BIN")) || bytes.Contains(out, []byte("vendor/mog")) {
		t.Fatalf("help must not name Mog env/vendor as a default engine:\n%s", out)
	}
	if err := run([]string{"--help"}); err != nil {
		t.Fatalf("run(--help) = %v", err)
	}
}

func TestRunVersion(t *testing.T) {
	if err := run([]string{"version"}); err != nil {
		t.Fatalf("run(version) = %v", err)
	}
}

func TestRunUnknown(t *testing.T) {
	err := run([]string{"nope"})
	if err == nil || !strings.Contains(err.Error(), "unknown command") {
		t.Fatalf("run(nope) = %v, want unknown command", err)
	}
}

func TestRunExcelSaveUsage(t *testing.T) {
	err := run([]string{"excel-save"})
	if err == nil || !strings.Contains(err.Error(), "usage:") {
		t.Fatalf("run(excel-save) = %v, want usage error", err)
	}
}

func TestRunExcelSaveOffWindows(t *testing.T) {
	if runtime.GOOS == "windows" || os.Getenv("GOOS_FORCE") == "windows" {
		t.Skip("windows host implements excel-save")
	}
	err := run([]string{"excel-save", "in.xlsx", "out.xlsx"})
	if err == nil {
		t.Fatal("expected error off Windows")
	}
	if !errors.Is(err, excel.ErrNotWindows) && !strings.Contains(err.Error(), "not implemented") {
		t.Fatalf("excel-save error = %v", err)
	}
}

func TestRunExcelRunUsage(t *testing.T) {
	err := run([]string{"excel-run"})
	if err == nil || !strings.Contains(err.Error(), "usage:") {
		t.Fatalf("run(excel-run) = %v, want usage error", err)
	}
	err = run([]string{"excel-run", "in.xlsx", "out.xlsx"})
	if err == nil || !strings.Contains(err.Error(), "usage:") {
		t.Fatalf("run(excel-run two args) = %v, want usage error", err)
	}
}

func TestRunExcelRunOffWindows(t *testing.T) {
	if runtime.GOOS == "windows" || os.Getenv("GOOS_FORCE") == "windows" {
		t.Skip("windows host implements excel-run")
	}
	err := run([]string{"excel-run", "in.xlsx", "script.js", "out.xlsx"})
	if err == nil {
		t.Fatal("expected error off Windows")
	}
	if !errors.Is(err, excel.ErrNotWindows) && !strings.Contains(err.Error(), "not implemented") {
		t.Fatalf("excel-run error = %v", err)
	}
}

func TestSidecarMetaScriptOnlyWhenRan(t *testing.T) {
	info := excel.HostInfo{ID: "excel-win", OS: "windows", ExcelVersion: "16.0", ExcelBuild: "1"}
	load := sidecarMeta(info, "init.xlsx", "")
	if load.Script != "" {
		t.Fatalf("load+save invented script %q", load.Script)
	}
	ran := sidecarMeta(info, filepath.Join("cases", "simple_set_a1", "init.xlsx"), filepath.Join("cases", "simple_set_a1", "script.js"))
	if ran.Script != "script.js" {
		t.Fatalf("script identity = %q", ran.Script)
	}
	dir := t.TempDir()
	xlsx := filepath.Join(dir, "golden.xlsx")
	if err := os.WriteFile(xlsx, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := golden.Write(xlsx, load); err != nil {
		t.Fatal(err)
	}
	raw, err := os.ReadFile(golden.PathFor(xlsx))
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Contains(raw, []byte(`"script"`)) {
		t.Fatalf("load+save sidecar must omit script:\n%s", raw)
	}
	if err := golden.Write(xlsx, ran); err != nil {
		t.Fatal(err)
	}
	got, err := golden.Read(xlsx)
	if err != nil {
		t.Fatal(err)
	}
	if got.Script != "script.js" {
		t.Fatalf("read script = %q", got.Script)
	}
}
