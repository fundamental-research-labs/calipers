package main

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestSpecsCountUniqueAndOfficeJS(t *testing.T) {
	specs := allSpecs()
	if n := len(specs); n < 120 || n > 160 {
		t.Fatalf("spec count %d, want 120-160", n)
	}
	seen := map[string]bool{}
	joined := strings.Builder{}
	for _, s := range specs {
		if s.name == "" {
			t.Fatal("empty spec name")
		}
		if seen[s.name] {
			t.Errorf("duplicate spec name %s", s.name)
		}
		seen[s.name] = true
		body := strings.TrimSpace(s.script)
		if body == "" {
			t.Errorf("%s: empty script", s.name)
			continue
		}
		if !strings.Contains(s.script, "Excel.run") {
			t.Errorf("%s: missing Excel.run", s.name)
		}
		if strings.Contains(s.script, "RAND(") || strings.Contains(s.script, "NOW(") || strings.Contains(s.script, "TODAY(") || strings.Contains(s.script, "RANDARRAY") {
			t.Errorf("%s: volatile RAND/NOW/TODAY/RANDARRAY is not golden-friendly", s.name)
		}
		joined.WriteString(s.script)
		joined.WriteByte('\n')
	}
	text := joined.String()
	for _, needle := range uncoveredMogMethodNeedles {
		if !strings.Contains(text, needle) {
			t.Errorf("specs missing %s", needle)
		}
	}
	for _, needle := range []string{
		"tables.add",
		"pivotTables.add",
		"charts.add",
		"FILTER(",
		"UNIQUE(",
		"SORT(",
		"XLOOKUP(",
		"SEQUENCE(",
		"format.font",
		"format.fill",
		"numberFormat",
		"conditionalFormats",
		"dataValidation",
		"autoFilter",
		"comments.add",
		"hyperlink",
		"names.add",
		"freezePanes",
		"tabColor",
		".sort.apply",
		".insert(",
		"for (let i = 1; i <= 500; i++)",
	} {
		if !strings.Contains(text, needle) {
			t.Errorf("specs missing %s", needle)
		}
	}
}

func TestCommittedOfficejsMatchesSpecs(t *testing.T) {
	root := filepath.Join("..", "..", "verification", "cases", "officejs")
	empty := filepath.Join("..", "..", "verification", "cases", "roundtrip", "empty", "init.xlsx")
	emptyBytes, err := os.ReadFile(empty)
	if err != nil {
		t.Fatal(err)
	}
	specs := allSpecs()
	for _, s := range specs {
		dir := filepath.Join(root, s.name)
		got, err := os.ReadFile(filepath.Join(dir, "script.js"))
		if err != nil {
			t.Errorf("%s: %v", s.name, err)
			continue
		}
		if strings.ReplaceAll(string(got), "\r\n", "\n") != s.script {
			t.Errorf("%s: committed script.js drifted from spec table", s.name)
		}
		init, err := os.ReadFile(filepath.Join(dir, "init.xlsx"))
		if err != nil {
			t.Errorf("%s: init: %v", s.name, err)
			continue
		}
		if string(init) != string(emptyBytes) {
			t.Errorf("%s: init.xlsx is not a copy of roundtrip/empty", s.name)
		}
		golden := filepath.Join(dir, "golden.xlsx")
		st, err := os.Stat(golden)
		if err != nil {
			if !os.IsNotExist(err) {
				t.Errorf("%s: golden: %v", s.name, err)
			}
			continue
		}
		if st.Size() == 0 {
			t.Errorf("%s: empty golden.xlsx", s.name)
		}
		if _, err := os.Stat(filepath.Join(dir, "config.json")); err != nil {
			t.Errorf("%s: has golden.xlsx but missing config.json", s.name)
		}
	}
}

// uncoveredMogMethodNeedles are Office.js methods mog implements that no
// pre-existing officejs/scratch/default script invoked. Each must appear in
// allSpecs() so generated cases cover the leftover mutators/navigators.
var uncoveredMogMethodNeedles = []string{
	"getCell(",
	"getRow(",
	"getColumn(",
	"getLastCell(",
	"getLastRow(",
	"getLastColumn(",
	"getResizedRange(",
	"getAbsoluteResizedRange(",
	"getRowsAbove(",
	"getRowsBelow(",
	"getColumnsBefore(",
	"getColumnsAfter(",
	"getBoundingRect(",
	"getIntersection(",
	".unmerge(",
	"format.fill.clear",
	".activate(",
	"doomed.delete(",
	".getNext(",
	".getPrevious(",
	".getFirst(",
	".getLast(",
	"worksheets.getItem(",
	"addFormulaLocal(",
	"named.delete(",
	"table.delete(",
	"convertToRange(",
	"table.resize(",
	"rows.getItemAt(",
	"deleteRows(",
	"deleteRowsAt(",
	`columns.getItem("Product").delete(`,
	"getHeaderRowRange(",
	"getDataBodyRange(",
	"getTotalRowRange(",
	"freezeAt(",
	".unfreeze(",
	"autoFilter.remove(",
	"clearCriteria(",
	"clearColumnCriteria(",
	".reapply(",
	"dataValidation.clear(",
}

func TestGenerateWritesInitAndScriptNoGolden(t *testing.T) {
	empty := filepath.Join("..", "..", "verification", "cases", "roundtrip", "empty", "init.xlsx")
	if _, err := os.Stat(empty); err != nil {
		t.Fatalf("empty init: %v", err)
	}
	dest := t.TempDir()
	n, err := generate(empty, dest)
	if err != nil {
		t.Fatal(err)
	}
	if n != len(allSpecs()) {
		t.Fatalf("wrote %d, want %d", n, len(allSpecs()))
	}
	entries, err := os.ReadDir(dest)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != n {
		t.Fatalf("dirs %d, want %d", len(entries), n)
	}
	emptyBytes, err := os.ReadFile(empty)
	if err != nil {
		t.Fatal(err)
	}
	sample := filepath.Join(dest, allSpecs()[0].name)
	got, err := os.ReadFile(filepath.Join(sample, "init.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	if string(got) != string(emptyBytes) {
		t.Fatal("generated init.xlsx must copy roundtrip/empty")
	}
	if _, err := os.Stat(filepath.Join(sample, "golden.xlsx")); err == nil {
		t.Fatal("generator must not write golden.xlsx")
	}
	if _, err := os.Stat(filepath.Join(sample, "config.json")); err == nil {
		t.Fatal("generator must not write config.json")
	}
}
