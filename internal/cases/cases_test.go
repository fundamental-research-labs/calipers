package cases

import (
	"bytes"
	"os"
	"path/filepath"
	"runtime"
	"testing"
)

func TestLoadMissingScriptIsLoadSave(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "tier_a_noscript", "pk", nil)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 {
		t.Fatalf("len=%d, want 1", len(got))
	}
	c := got[0]
	if c.ID != "tier_a_noscript" || c.Tier != TierA {
		t.Fatalf("id/tier = %s %s", c.ID, c.Tier)
	}
	if c.RunScript() || c.ScriptPath != "" {
		t.Fatalf("missing script should be load+save, got ScriptPath=%q RunScript=%v", c.ScriptPath, c.RunScript())
	}
	if filepath.Base(c.InitPath) != InitFile {
		t.Fatalf("InitPath=%q", c.InitPath)
	}
	if filepath.Base(c.GoldenPath) != GoldenFile {
		t.Fatalf("GoldenPath=%q", c.GoldenPath)
	}
}

func TestLoadEmptyScriptIsLoadSave(t *testing.T) {
	root := t.TempDir()
	empty := ""
	writeCase(t, root, "tier_a_emptyjs", "pk", &empty)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 {
		t.Fatalf("len=%d, want 1", len(got))
	}
	c := got[0]
	if c.RunScript() || c.ScriptPath != "" {
		t.Fatalf("empty script should be load+save, got ScriptPath=%q RunScript=%v", c.ScriptPath, c.RunScript())
	}
	if _, err := os.Stat(filepath.Join(c.Dir, ScriptFile)); err != nil {
		t.Fatalf("empty script.js should still exist on disk: %v", err)
	}
}

func TestLoadWhitespaceScriptIsLoadSave(t *testing.T) {
	root := t.TempDir()
	ws := "  \n\t\n"
	writeCase(t, root, "tier_b_ws", "pk", &ws)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].Tier != TierB {
		t.Fatalf("got %+v", got)
	}
	if got[0].RunScript() {
		t.Fatal("whitespace-only script should skip execution")
	}
}

func TestLoadNonEmptyScript(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async () => { await Excel.run(); });\n"
	writeCase(t, root, "tier_a_js", "pk", &js)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 {
		t.Fatalf("len=%d", len(got))
	}
	c := got[0]
	if !c.RunScript() {
		t.Fatal("non-empty script should run")
	}
	if filepath.Base(c.ScriptPath) != ScriptFile {
		t.Fatalf("ScriptPath=%q", c.ScriptPath)
	}
	data, err := os.ReadFile(c.ScriptPath)
	if err != nil {
		t.Fatal(err)
	}
	if string(data) != js {
		t.Fatalf("script body = %q, want %q", data, js)
	}
}

func TestDefaultPassExcludesTierC(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "tier_a_one", "pk", nil)
	writeCase(t, root, "tier_b_two", "pk", nil)
	writeCase(t, root, "tier_c_bomb", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(all) != 3 {
		t.Fatalf("len(all)=%d", len(all))
	}
	pass := DefaultPass(all)
	if len(pass) != 2 {
		t.Fatalf("default pass len=%d, want 2", len(pass))
	}
	for _, c := range pass {
		if c.Tier == TierC {
			t.Fatalf("default pass included %s", c.ID)
		}
	}
	ids := map[string]bool{pass[0].ID: true, pass[1].ID: true}
	if !ids["tier_a_one"] || !ids["tier_b_two"] {
		t.Fatalf("default pass ids = %v", ids)
	}
}

func TestLoadIgnoresNonCaseDirs(t *testing.T) {
	root := t.TempDir()
	if err := os.Mkdir(filepath.Join(root, "not_a_case"), 0o755); err != nil {
		t.Fatal(err)
	}
	writeCase(t, root, "tier_a_ok", "pk", nil)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "tier_a_ok" {
		t.Fatalf("got %+v", got)
	}
}

func TestLoadRequiresInit(t *testing.T) {
	root := t.TempDir()
	if err := os.Mkdir(filepath.Join(root, "tier_a_missing_init"), 0o755); err != nil {
		t.Fatal(err)
	}
	_, err := Load(root)
	if err == nil {
		t.Fatal("expected error for tier dir without init.xlsx")
	}
}

func TestLoadRealCorpus(t *testing.T) {
	root := repoCasesDir(t)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(all) != 91 {
		t.Fatalf("real corpus: got %d cases, want 91", len(all))
	}
	var nA, nB, nC int
	for _, c := range all {
		switch c.Tier {
		case TierA:
			nA++
		case TierB:
			nB++
		case TierC:
			nC++
		default:
			t.Errorf("%s: unexpected tier %q", c.ID, c.Tier)
		}
		if c.ID != "tier_a_simple" && (c.RunScript() || c.ScriptPath != "") {
			t.Errorf("%s: unexpected script %q (only tier_a_simple is scripted)", c.ID, c.ScriptPath)
		}
		if _, err := os.Stat(c.InitPath); err != nil {
			t.Errorf("%s: init: %v", c.ID, err)
		}
		if filepath.Base(c.InitPath) != InitFile {
			t.Errorf("%s: init must be %s, got %s", c.ID, InitFile, c.InitPath)
		}
		if _, err := os.Stat(c.GoldenPath); err == nil {
			t.Errorf("%s: golden xlsx must not live among case inputs yet: %s", c.ID, c.GoldenPath)
		}
	}
	if nA != 57 || nB != 26 || nC != 8 {
		t.Fatalf("tier counts a=%d b=%d c=%d, want 57/26/8", nA, nB, nC)
	}

	var scripted []Case
	for _, c := range all {
		if c.RunScript() {
			scripted = append(scripted, c)
		}
	}
	if len(scripted) != 1 {
		t.Fatalf("scripted cases: got %d, want 1", len(scripted))
	}
	s := scripted[0]
	if s.ID != "tier_a_simple" {
		t.Fatalf("scripted case = %s, want tier_a_simple", s.ID)
	}
	if filepath.Base(s.ScriptPath) != ScriptFile {
		t.Fatalf("ScriptPath=%q", s.ScriptPath)
	}
	body, err := os.ReadFile(s.ScriptPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(bytes.TrimSpace(body)) == 0 {
		t.Fatal("tier_a_simple script.js is empty; discovery would skip it")
	}
	if !bytes.Contains(body, []byte("Excel.run")) {
		t.Fatalf("tier_a_simple script.js is not Office.js Excel.run:\n%s", body)
	}

	pass := DefaultPass(all)
	if len(pass) != nA+nB {
		t.Fatalf("default pass len=%d, want %d (exclude tier_c)", len(pass), nA+nB)
	}
	for _, c := range pass {
		if c.Tier == TierC {
			t.Errorf("default pass included hostile %s", c.ID)
		}
	}
	if len(pass) != len(all)-nC {
		t.Fatalf("default pass %d vs all-c %d", len(pass), len(all)-nC)
	}
}

func writeCase(t *testing.T, root, id, init string, script *string) {
	t.Helper()
	dir := filepath.Join(root, id)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, InitFile), []byte(init), 0o644); err != nil {
		t.Fatal(err)
	}
	if script != nil {
		if err := os.WriteFile(filepath.Join(dir, ScriptFile), []byte(*script), 0o644); err != nil {
			t.Fatal(err)
		}
	}
}

func repoCasesDir(t *testing.T) string {
	t.Helper()
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("runtime.Caller failed")
	}
	dir := filepath.Join(filepath.Dir(file), "..", "..", DirName)
	abs, err := filepath.Abs(dir)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(abs); err != nil {
		t.Fatalf("corpus %s: %v", abs, err)
	}
	return abs
}
