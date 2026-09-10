package xlsxmodel

import (
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"testing"
)

// roundtripInitGoldenMismatches are roundtrip cases whose Excel load+save
// golden is not semantically equal to init.xlsx. Inits are themselves
// Windows Excel exports, so load+save is a no-op except for volatiles
// recalculated at save time.
var roundtripInitGoldenMismatches = map[string]string{
	"formula_stress_test":     "volatile NOW/TODAY/RAND cached values",
	"formulas_datetime":       "volatile NOW/TODAY serials",
	"formulas_dynamic_arrays": "volatile RAND values",
	"formulas_information":    "CELL filename depends on path",
}

func TestCompareFilesRoundtripInitVsGolden(t *testing.T) {
	root := roundtripCasesDir(t)
	entries, err := os.ReadDir(root)
	if err != nil {
		t.Fatal(err)
	}
	var names []string
	for _, e := range entries {
		if e.IsDir() {
			names = append(names, e.Name())
		}
	}
	sort.Strings(names)
	if len(names) == 0 {
		t.Fatal("no roundtrip cases found")
	}

	seen := map[string]bool{}
	for _, name := range names {
		seen[name] = true
		t.Run(name, func(t *testing.T) {
			initPath := filepath.Join(root, name, "init.xlsx")
			goldenPath := filepath.Join(root, name, "golden.xlsx")
			got, err := CompareFiles(initPath, goldenPath)
			if err != nil {
				t.Fatal(err)
			}
			reason, wantDiff := roundtripInitGoldenMismatches[name]
			if wantDiff {
				if got.Equal {
					t.Fatalf("expected mismatch (%s); remove from roundtripInitGoldenMismatches", reason)
				}
				return
			}
			if !got.Equal {
				t.Fatalf("init vs golden must match, got %v", got.Diffs)
			}
		})
	}
	for name := range roundtripInitGoldenMismatches {
		if !seen[name] {
			t.Errorf("roundtripInitGoldenMismatches %q is not a roundtrip case", name)
		}
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
