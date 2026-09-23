package main

import (
	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
	"github.com/fundamental-research-labs/calipers/internal/golden"
	"os"
	"path/filepath"
	"testing"
)

type recalcHost struct {
	fakeEngine
	recalcs int
}

func (h *recalcHost) RecalculateOpenSave(in, out string) error { h.recalcs++; return copyFile(in, out) }
func (h *recalcHost) Available() bool                          { return true }
func (h *recalcHost) Info() (excel.HostInfo, error) {
	return excel.HostInfo{ID: excel.HostID, OS: "windows", ExcelVersion: "test"}, nil
}

func TestRecalculationAcrossGoldenVerifyAndMeasurement(t *testing.T) {
	root, out := setupVerifyDir(t)
	writeVerifyCase(t, root, "recalc", "hello", nil)
	dir := filepath.Join(root, "recalc")
	if err := os.WriteFile(filepath.Join(dir, "config.json"), []byte(`{"recalculate":true}`), 0644); err != nil {
		t.Fatal(err)
	}
	all, err := cases.Load(root)
	if err != nil {
		t.Fatal(err)
	}
	c := all[0]
	h := &recalcHost{}
	if err := generateCaseGolden(h, c); err != nil {
		t.Fatal(err)
	}
	meta, err := golden.Read(c.GoldenPath)
	if err != nil || !meta.Recalculate {
		t.Fatalf("meta=%+v err=%v", meta, err)
	}
	if result := verifyOne(h, c, out, false); result.Status != "pass" {
		t.Fatalf("%+v", result)
	}
	if _, err := measureExcel(h, c, filepath.Join(out, "measured.xlsx")); err != nil {
		t.Fatal(err)
	}
	if h.recalcs != 3 || len(h.opens) != 0 {
		t.Fatalf("recalcs=%d opens=%v", h.recalcs, h.opens)
	}
	c.Recalculate = false
	if err := executeCase(h, c, filepath.Join(out, "plain.xlsx")); err != nil {
		t.Fatal(err)
	}
	if len(h.opens) != 1 {
		t.Fatal("recalculation leaked into ordinary case")
	}
	c.Recalculate = true
	if err := executeCase(&fakeEngine{}, c, filepath.Join(out, "unsupported.xlsx")); err == nil {
		t.Fatal("silently skipped recalculation")
	}
}
