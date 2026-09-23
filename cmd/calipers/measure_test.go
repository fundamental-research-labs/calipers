package main

import (
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
	"time"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
)

func TestRunMeasureBudgetsHelp(t *testing.T) {
	out, err := captureStdout(t, func() error {
		return run([]string{"measure-budgets", "-h"})
	})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(out, "config.json") || !strings.Contains(out, "Windows") {
		t.Fatalf("help must mention config.json and Windows:\n%s", out)
	}
	if !strings.Contains(out, "maxPeakMemoryBytes") || !strings.Contains(out, "maxDurationMs") {
		t.Fatalf("help must name budget fields:\n%s", out)
	}
}

func TestRunMeasureBudgetsUsage(t *testing.T) {
	err := run([]string{"measure-budgets", "a", "b", "c"})
	if err == nil || !strings.Contains(err.Error(), "usage:") {
		t.Fatalf("extra args = %v, want usage error", err)
	}
}

func TestRunMeasureBudgetsEmptyDir(t *testing.T) {
	dir := t.TempDir()
	if err := run([]string{"measure-budgets", "--cases-dir", dir}); err != nil {
		t.Fatalf("empty corpus should not require Excel: %v", err)
	}
}

func TestRunMeasureBudgetsSkipsExistingConfig(t *testing.T) {
	root := t.TempDir()
	writeCaseDir(t, root, "capped", nil)
	if err := cases.WriteBudget(filepath.Join(root, "capped"), cases.Budget{MaxDurationMs: 10}); err != nil {
		t.Fatal(err)
	}
	if err := run([]string{"measure-budgets", "--cases-dir", root}); err != nil {
		t.Fatalf("nothing to measure should not require Excel: %v", err)
	}
}

func TestRunMeasureBudgetsExcelOffWindows(t *testing.T) {
	if runtime.GOOS == "windows" || os.Getenv("GOOS_FORCE") == "windows" {
		t.Skip("windows host implements measure-budgets")
	}
	root := t.TempDir()
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeCaseDir(t, root, "pending", &js)
	err := run([]string{"measure-budgets", "--engine", "excel", "--cases-dir", root})
	if err == nil {
		t.Fatal("expected error off Windows")
	}
	if !errors.Is(err, excel.ErrNotWindows) && !strings.Contains(err.Error(), "Windows") {
		t.Fatalf("measure-budgets error = %v", err)
	}
}

func TestApplyMargin(t *testing.T) {
	if got := applyMargin(100, 1.5, 10); got != 150 {
		t.Fatalf("100*1.5 = %d, want 150", got)
	}
	if got := applyMargin(10, 1.5, 20); got != 30 {
		t.Fatalf("floor should win: got %d", got)
	}
	if got := applyMargin(0, 1.5, 20); got != 0 {
		t.Fatalf("zero stays zero: %d", got)
	}
}

func TestBudgetFromSample(t *testing.T) {
	b := budgetFromSample(memSample{peakBytes: 100 << 20, duration: 2 * time.Second}, 1.5)
	if b.MaxPeakMemoryBytes < 150<<20 {
		t.Fatalf("memory budget %d, want at least 1.5x", b.MaxPeakMemoryBytes)
	}
	if b.MaxDurationMs < 3000 {
		t.Fatalf("duration budget %d, want at least 1.5x 2000ms", b.MaxDurationMs)
	}
}

func TestBinaryBudgetRunnerHonorsCaseRecalculation(t *testing.T) {
	if runtime.GOOS == "windows" {
		t.Skip("shell fixture")
	}
	dir := t.TempDir()
	bin := filepath.Join(dir, "engine")
	// A plain save or a misplaced flag must fail, so this covers the real child argv.
	script := "#!/bin/sh\n[ \"$1\" = save ] && [ \"$2\" = --recalculate ] || exit 2\ncp \"$3\" \"$4\"\n"
	if err := os.WriteFile(bin, []byte(script), 0755); err != nil {
		t.Fatal(err)
	}
	in := filepath.Join(dir, "init.xlsx")
	if err := os.WriteFile(in, []byte("pk"), 0644); err != nil {
		t.Fatal(err)
	}
	run, err := budgetRunner(bin)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := run(cases.Case{InitPath: in, Recalculate: true}); err != nil {
		t.Fatal(err)
	}
}
