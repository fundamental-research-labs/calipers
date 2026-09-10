package main

import (
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
)

func TestRunUnavailableHost(t *testing.T) {
	err := run(&fakeHost{available: false}, []string{t.TempDir()})
	if !errors.Is(err, excel.ErrNotWindows) {
		t.Fatalf("error = %v, want ErrNotWindows", err)
	}
}

func TestRunNewHostOffWindows(t *testing.T) {
	if runtime.GOOS == "windows" {
		t.Skip("windows host implements excel-export-inits")
	}
	err := run(excel.NewHost(), []string{t.TempDir()})
	if !errors.Is(err, excel.ErrNotWindows) {
		t.Fatalf("error = %v, want ErrNotWindows", err)
	}
}

func TestRunUsage(t *testing.T) {
	err := run(&fakeHost{available: true}, []string{"a", "b"})
	if err == nil || !strings.Contains(err.Error(), "usage:") {
		t.Fatalf("error = %v, want usage", err)
	}
}

func TestRewriteLiveExcelOneMacFixture(t *testing.T) {
	h := excel.NewHost()
	if !h.Available() {
		t.Skip("Excel.Application ProgID not registered")
	}
	src := filepath.Join(repoRoot(t), cases.DirName, "roundtrip", "simple", cases.InitFile)
	orig, err := os.ReadFile(src)
	if err != nil {
		t.Fatal(err)
	}
	root := t.TempDir()
	initPath := writeCaseInit(t, root, "roundtrip", "simple", orig)
	loaded, err := cases.Load(root)
	if err != nil {
		t.Fatal(err)
	}

	var buf strings.Builder
	if err := rewriteInits(h, loaded, &buf); err != nil {
		t.Fatalf("rewrite: %v\n%s", err, buf.String())
	}
	if err := excel.CheckWindowsExcel16Export(initPath); err != nil {
		t.Fatal(err)
	}
	app, ver, err := excel.ReadAppProperties(initPath)
	if err != nil {
		t.Fatal(err)
	}
	if app != excel.WindowsExcelApplication || !strings.HasPrefix(ver, "16.") {
		t.Fatalf("exported Application=%q AppVersion=%q", app, ver)
	}
}

func repoRoot(t *testing.T) string {
	t.Helper()
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("runtime.Caller failed")
	}
	abs, err := filepath.Abs(filepath.Join(filepath.Dir(file), "..", ".."))
	if err != nil {
		t.Fatal(err)
	}
	return abs
}
