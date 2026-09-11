package xlsxmodel_test

import (
	"path/filepath"
	"runtime"
	"testing"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/xlsxmodel"
)

func TestCompareFilesRoundtripInitVsGolden(t *testing.T) {
	root := roundtripCasesDir(t)
	all, err := cases.Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(all) == 0 {
		t.Fatal("no roundtrip cases found")
	}
	for _, c := range all {
		c := c
		t.Run(c.Name, func(t *testing.T) {
			got, err := xlsxmodel.CompareFiles(c.InitPath, c.GoldenPath, c.Compare)
			if err != nil {
				t.Fatal(err)
			}
			if !got.Equal {
				t.Fatalf("init vs golden must match under resolved config.json, got %v", got.Diffs)
			}
		})
	}
}

func TestCompareFilesRoundtripNonExceptedMismatchStillFails(t *testing.T) {
	root := roundtripCasesDir(t)
	all, err := cases.Load(root)
	if err != nil {
		t.Fatal(err)
	}
	var datetime cases.Case
	for _, c := range all {
		if c.Name == "formulas_datetime" {
			datetime = c
			break
		}
	}
	if datetime.Name == "" {
		t.Fatal("missing formulas_datetime")
	}
	got, err := xlsxmodel.CompareFiles(datetime.InitPath, filepath.Join(root, "simple", "golden.xlsx"), datetime.Compare)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("comparing datetime init to an unrelated golden must still fail")
	}
}

func roundtripCasesDir(t *testing.T) string {
	t.Helper()
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("runtime.Caller failed")
	}
	return filepath.Join(filepath.Dir(file), "..", "..", "verification", "cases", "roundtrip")
}
