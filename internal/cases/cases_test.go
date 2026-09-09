package cases

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"
	"testing"

	"github.com/fundamental-research-labs/calipers/internal/excel"
	"github.com/fundamental-research-labs/calipers/internal/golden"
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
	if c.ID != "tier_a_noscript" || c.Name != "tier_a_noscript" || c.Tier != TierA {
		t.Fatalf("id/name/tier = %s %s %s", c.ID, c.Name, c.Tier)
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

func TestOpenSavePassSkipsScriptedAndTierC(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async (ctx) => {});"
	writeCase(t, root, "tier_a_one", "pk", nil)
	writeCase(t, root, "tier_a_js", "pk", &js)
	writeCase(t, root, "tier_b_two", "pk", nil)
	writeCase(t, root, "tier_c_bomb", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	pass := OpenSavePass(all)
	if len(pass) != 2 {
		t.Fatalf("open+save pass len=%d, want 2", len(pass))
	}
	ids := map[string]bool{}
	for _, c := range pass {
		ids[c.ID] = true
		if c.Tier == TierC {
			t.Fatalf("open+save pass included hostile %s", c.ID)
		}
		if c.RunScript() {
			t.Fatalf("open+save pass included scripted %s", c.ID)
		}
	}
	if !ids["tier_a_one"] || !ids["tier_b_two"] {
		t.Fatalf("open+save pass ids = %v", ids)
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

func TestSelectDefaultPassWhenNoIDs(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "tier_a_one", "pk", nil)
	writeCase(t, root, "tier_b_two", "pk", nil)
	writeCase(t, root, "tier_c_bomb", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	got, err := Select(all, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 2 {
		t.Fatalf("len=%d, want default pass 2", len(got))
	}
	for _, c := range got {
		if c.Tier == TierC {
			t.Fatalf("default select included %s", c.ID)
		}
	}
}

func TestSelectByIDIncludesTierCAndOrder(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "tier_a_one", "pk", nil)
	writeCase(t, root, "tier_c_bomb", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	got, err := Select(all, []string{"tier_c_bomb", "tier_a_one"})
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 2 || got[0].ID != "tier_c_bomb" || got[1].ID != "tier_a_one" {
		t.Fatalf("got %+v", got)
	}
	_, err = Select(all, []string{"tier_a_nope"})
	if err == nil || !bytes.Contains([]byte(err.Error()), []byte("unknown case")) {
		t.Fatalf("want unknown case, got %v", err)
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

func TestLoadNestedSuites(t *testing.T) {
	root := t.TempDir()
	writeCase(t, filepath.Join(root, "roundtrip"), "tier_a_rt", "pk", nil)
	writeCase(t, filepath.Join(root, "default"), "tier_a_other", "pk", nil)
	writeCase(t, filepath.Join(root, "default"), "tier_c_bomb", "pk", nil)
	if err := os.Mkdir(filepath.Join(root, "empty_suite"), 0o755); err != nil {
		t.Fatal(err)
	}

	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(corpus.Cases) != 3 {
		t.Fatalf("len=%d, want 3", len(corpus.Cases))
	}
	wantSuites := []string{"default", "empty_suite", "roundtrip"}
	if strings.Join(corpus.Suites, ",") != strings.Join(wantSuites, ",") {
		t.Fatalf("suites=%v, want %v", corpus.Suites, wantSuites)
	}
	byID := map[string]Case{}
	for _, c := range corpus.Cases {
		byID[c.ID] = c
		if c.Suite == "" {
			t.Errorf("%s: suite should be set in nested layout", c.ID)
		}
		if c.ID != c.Suite+"/"+c.Name {
			t.Errorf("%s: ID must be suite/name, Name=%q Suite=%q", c.ID, c.Name, c.Suite)
		}
		if !strings.Contains(c.Dir, filepath.FromSlash(c.ID)) {
			t.Errorf("%s: Dir=%q does not include id", c.ID, c.Dir)
		}
	}
	if byID["roundtrip/tier_a_rt"].Suite != "roundtrip" {
		t.Fatalf("roundtrip/tier_a_rt suite=%q", byID["roundtrip/tier_a_rt"].Suite)
	}
	if byID["default/tier_a_other"].Suite != "default" || byID["default/tier_c_bomb"].Suite != "default" {
		t.Fatalf("default-suite cases: %+v", byID)
	}

	got, err := FilterSuite(corpus.Cases, "roundtrip", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "roundtrip/tier_a_rt" {
		t.Fatalf("roundtrip filter = %+v", got)
	}
	empty, err := FilterSuite(corpus.Cases, "empty_suite", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if len(empty) != 0 {
		t.Fatalf("empty suite = %+v", empty)
	}
	_, err = FilterSuite(corpus.Cases, "nope", corpus.Suites)
	if err == nil || !bytes.Contains([]byte(err.Error()), []byte("unknown suite")) {
		t.Fatalf("want unknown suite, got %v", err)
	}

	pass := DefaultPass(corpus.Cases)
	if len(pass) != 2 {
		t.Fatalf("default pass across suites len=%d, want 2", len(pass))
	}
}

func TestSelectQualifiedID(t *testing.T) {
	root := t.TempDir()
	writeCase(t, filepath.Join(root, "roundtrip"), "tier_a_one", "pk", nil)
	writeCase(t, filepath.Join(root, "default"), "tier_c_bomb", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	got, err := Select(all, []string{"roundtrip/tier_a_one", "default/tier_c_bomb"})
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 2 || got[0].ID != "roundtrip/tier_a_one" || got[1].ID != "default/tier_c_bomb" {
		t.Fatalf("got %+v", got)
	}
	bare, err := Select(all, []string{"tier_a_one"})
	if err != nil {
		t.Fatal(err)
	}
	if len(bare) != 1 || bare[0].ID != "roundtrip/tier_a_one" {
		t.Fatalf("unique bare name = %+v", bare)
	}
}

func TestSelectAmbiguousID(t *testing.T) {
	root := t.TempDir()
	writeCase(t, filepath.Join(root, "roundtrip"), "tier_a_one", "pk", nil)
	writeCase(t, filepath.Join(root, "default"), "tier_a_one", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	_, err = Select(all, []string{"tier_a_one"})
	if err == nil || !bytes.Contains([]byte(err.Error()), []byte("ambiguous case")) {
		t.Fatalf("want ambiguous case, got %v", err)
	}
	got, err := Select(all, []string{"default/tier_a_one"})
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "default/tier_a_one" {
		t.Fatalf("got %+v", got)
	}
}

func TestGoldenComparePassSkipsMissingGolden(t *testing.T) {
	root := t.TempDir()
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeCase(t, filepath.Join(root, "roundtrip"), "tier_a_rt", "pk", nil)
	writeCase(t, filepath.Join(root, "scratch"), "text", "pk", &js)
	writeCase(t, filepath.Join(root, "scratch"), "tier_c_bomb", "pk", &js)

	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	var rt *Case
	for i := range all {
		if all[i].ID == "roundtrip/tier_a_rt" {
			rt = &all[i]
			break
		}
	}
	if rt == nil {
		t.Fatal("missing roundtrip/tier_a_rt")
	}
	if err := os.WriteFile(rt.GoldenPath, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	def := DefaultPass(all)
	if len(def) != 2 {
		t.Fatalf("DefaultPass len=%d, want 2 (roundtrip a + unprefixed case without golden)", len(def))
	}
	got := GoldenComparePass(all)
	if len(got) != 1 || got[0].ID != "roundtrip/tier_a_rt" {
		t.Fatalf("GoldenComparePass = %+v, want only the case with a golden", got)
	}
}

func TestLoadFlatLayoutLeavesSuiteEmpty(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "tier_a_flat", "pk", nil)
	if err := os.Mkdir(filepath.Join(root, "not_a_suite"), 0o755); err != nil {
		t.Fatal(err)
	}
	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(corpus.Cases) != 1 || corpus.Cases[0].Suite != "" || corpus.Cases[0].ID != "tier_a_flat" || corpus.Cases[0].Name != "tier_a_flat" {
		t.Fatalf("got %+v", corpus)
	}
	if len(corpus.Suites) != 0 {
		t.Fatalf("flat layout should not report nested suites, got %v", corpus.Suites)
	}
}

func TestLoadRealCorpus(t *testing.T) {
	root := repoCasesDir(t)
	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	all := corpus.Cases
	if len(all) != 105 {
		t.Fatalf("real corpus: got %d cases, want 105", len(all))
	}
	wantSuites := "default,roundtrip,scratch"
	if strings.Join(corpus.Suites, ",") != wantSuites {
		t.Fatalf("real corpus suites=%v, want %s", corpus.Suites, wantSuites)
	}
	entries, err := os.ReadDir(root)
	if err != nil {
		t.Fatal(err)
	}
	for _, e := range entries {
		if e.IsDir() {
			if _, ok := parseTier(e.Name()); ok {
				t.Errorf("case %s must live under a suite directory, not %s", e.Name(), DirName)
			}
		}
	}
	var nA, nB, nC, nRT, nDef, nScratch int
	for _, c := range all {
		switch c.Tier {
		case TierA:
			nA++
		case TierB:
			nB++
		case TierC:
			nC++
		case "":
			if c.Suite != "scratch" {
				t.Errorf("%s: unexpected empty tier", c.ID)
			}
		default:
			t.Errorf("%s: unexpected tier %q", c.ID, c.Tier)
		}
		switch c.Suite {
		case "roundtrip":
			nRT++
			if c.RunScript() {
				t.Errorf("%s: roundtrip suite must stay load+save, got script %q", c.ID, c.ScriptPath)
			}
		case "default":
			nDef++
		case "scratch":
			nScratch++
			if !c.RunScript() {
				t.Errorf("%s: scratch case must be scripted", c.ID)
			}
			if c.Tier != "" || strings.HasPrefix(c.Name, "tier_") {
				t.Errorf("%s: scratch case must not use a tier_ prefix (tier=%q name=%q)", c.ID, c.Tier, c.Name)
			}
		default:
			t.Errorf("%s: unexpected suite %q", c.ID, c.Suite)
		}
		if c.ID != c.Suite+"/"+c.Name {
			t.Errorf("%s: ID must be suite/name, Name=%q Suite=%q", c.ID, c.Name, c.Suite)
		}
		if c.Suite != "scratch" && c.Name != "tier_a_simple_set_a1" && (c.RunScript() || c.ScriptPath != "") {
			t.Errorf("%s: unexpected script %q (only default/tier_a_simple_set_a1 and scratch/* are scripted)", c.ID, c.ScriptPath)
		}
		if _, err := os.Stat(c.InitPath); err != nil {
			t.Errorf("%s: init: %v", c.ID, err)
		}
		if filepath.Base(c.InitPath) != InitFile {
			t.Errorf("%s: init must be %s, got %s", c.ID, InitFile, c.InitPath)
		}
		if filepath.Base(c.GoldenPath) != GoldenFile {
			t.Errorf("%s: golden dest must be %s, got %s", c.ID, GoldenFile, c.GoldenPath)
		}
		wantDir := filepath.Join(root, filepath.FromSlash(c.ID))
		if c.Dir != wantDir {
			t.Errorf("%s: Dir=%q, want %q", c.ID, c.Dir, wantDir)
		}
	}
	if nA != 57 || nB != 25 || nC != 8 {
		t.Fatalf("tier counts a=%d b=%d c=%d, want 57/25/8", nA, nB, nC)
	}
	if nRT != 81 || nDef != 9 || nScratch != 15 {
		t.Fatalf("suite counts roundtrip=%d default=%d scratch=%d, want 81/9/15", nRT, nDef, nScratch)
	}

	var scripted []Case
	var simpleSetA1 *Case
	for i := range all {
		c := &all[i]
		if c.RunScript() {
			scripted = append(scripted, *c)
		}
		if c.ID == "default/tier_a_simple_set_a1" {
			simpleSetA1 = c
		}
	}
	if len(scripted) != 1+nScratch {
		t.Fatalf("scripted cases: got %d, want %d (default/tier_a_simple_set_a1 + scratch)", len(scripted), 1+nScratch)
	}
	if simpleSetA1 == nil || !simpleSetA1.RunScript() || simpleSetA1.Suite != "default" || simpleSetA1.Name != "tier_a_simple_set_a1" {
		t.Fatalf("scripted case default/tier_a_simple_set_a1 missing or not runnable")
	}
	s := *simpleSetA1
	if filepath.Base(s.ScriptPath) != ScriptFile {
		t.Fatalf("ScriptPath=%q", s.ScriptPath)
	}
	body, err := os.ReadFile(s.ScriptPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(bytes.TrimSpace(body)) == 0 {
		t.Fatal("tier_a_simple_set_a1 script.js is empty; discovery would skip it")
	}
	if !bytes.Contains(body, []byte("Excel.run")) {
		t.Fatalf("tier_a_simple_set_a1 script.js is not Office.js Excel.run:\n%s", body)
	}
	var simple *Case
	for i := range all {
		if all[i].ID == "roundtrip/tier_a_simple" {
			simple = &all[i]
			break
		}
	}
	if simple == nil {
		t.Fatal("missing load+save case roundtrip/tier_a_simple")
	}
	named, err := Select(all, []string{"roundtrip/tier_a_simple", "default/tier_a_simple_set_a1"})
	if err != nil {
		t.Fatal(err)
	}
	if len(named) != 2 || named[0].ID != "roundtrip/tier_a_simple" || named[1].ID != "default/tier_a_simple_set_a1" {
		t.Fatalf("Select by suite/name = %+v", named)
	}
	if simple.Suite != "roundtrip" {
		t.Fatalf("tier_a_simple suite=%q, want %s", simple.Suite, "roundtrip")
	}
	if simple.RunScript() || simple.ScriptPath != "" {
		t.Fatalf("tier_a_simple must stay load+save, got ScriptPath=%q", simple.ScriptPath)
	}
	simpleInit, err := os.ReadFile(simple.InitPath)
	if err != nil {
		t.Fatal(err)
	}
	scriptedInit, err := os.ReadFile(s.InitPath)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(simpleInit, scriptedInit) {
		t.Fatal("tier_a_simple_set_a1 init.xlsx must be a copy of tier_a_simple")
	}

	pass := DefaultPass(all)
	if len(pass) != nA+nB+nScratch {
		t.Fatalf("default pass len=%d, want %d (exclude tier_c, include unprefixed scratch)", len(pass), nA+nB+nScratch)
	}
	for _, c := range pass {
		if c.Tier == TierC {
			t.Errorf("default pass included hostile %s", c.ID)
		}
	}
	if len(pass) != len(all)-nC {
		t.Fatalf("default pass %d vs all-c %d", len(pass), len(all)-nC)
	}

	goldenPass := GoldenComparePass(all)
	if len(goldenPass) != nA+nB+nScratch {
		t.Fatalf("golden-compare pass len=%d, want %d (default pass with committed goldens)", len(goldenPass), nA+nB+nScratch)
	}
	for _, c := range goldenPass {
		st, err := os.Stat(c.GoldenPath)
		if err != nil || st.Size() == 0 {
			t.Errorf("%s: GoldenComparePass included a case with no golden", c.ID)
		}
		if c.Tier == TierC {
			t.Errorf("golden-compare pass included hostile %s", c.ID)
		}
	}

	openSave := OpenSavePass(all)
	if len(openSave) != nA+nB+nScratch-len(scripted) {
		t.Fatalf("open+save pass len=%d, want %d (default minus scripted)", len(openSave), nA+nB+nScratch-len(scripted))
	}
	for _, c := range openSave {
		if c.RunScript() {
			t.Errorf("open+save pass included scripted %s", c.ID)
		}
		if c.Tier == TierC {
			t.Errorf("open+save pass included hostile %s", c.ID)
		}
		if c.Suite == "scratch" {
			t.Errorf("open+save pass included scratch %s", c.ID)
		}
	}

	bySuite, err := FilterSuite(all, "scratch", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if len(bySuite) != nScratch {
		t.Fatalf("FilterSuite scratch len=%d, want %d", len(bySuite), nScratch)
	}
	for _, c := range bySuite {
		if c.Suite != "scratch" || !strings.HasPrefix(c.ID, "scratch/") {
			t.Errorf("FilterSuite scratch returned %s", c.ID)
		}
	}
}

func TestCommittedOpenSaveGoldens(t *testing.T) {
	root := repoCasesDir(t)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}

	var first golden.Meta
	var firstID string
	for i, c := range OpenSavePass(all) {
		st, err := os.Stat(c.GoldenPath)
		if err != nil {
			t.Errorf("%s: missing golden.xlsx", c.ID)
			continue
		}
		if st.Size() == 0 {
			t.Errorf("%s: golden.xlsx is empty", c.ID)
		}
		m, err := golden.Read(c.GoldenPath)
		if err != nil {
			t.Errorf("%s: %v", c.ID, err)
			continue
		}
		if m.Host != excel.HostID {
			t.Errorf("%s: host %q, want %s", c.ID, m.Host, excel.HostID)
		}
		if m.Script != "" {
			t.Errorf("%s: load+save golden must omit script, got %q", c.ID, m.Script)
		}
		if m.Input != InitFile {
			t.Errorf("%s: input %q, want %s", c.ID, m.Input, InitFile)
		}
		if m.ExcelVersion == "" || m.ExcelBuild == "" {
			t.Errorf("%s: missing Excel version/build in meta", c.ID)
		}
		if i == 0 {
			first = m
			firstID = c.ID
			continue
		}
		if m.Host != first.Host || m.ExcelVersion != first.ExcelVersion || m.ExcelBuild != first.ExcelBuild {
			t.Errorf("mixed excel goldens: %s is host=%s version=%s build=%s; %s is host=%s version=%s build=%s",
				c.ID, m.Host, m.ExcelVersion, m.ExcelBuild,
				firstID, first.Host, first.ExcelVersion, first.ExcelBuild)
		}
	}

	for _, c := range all {
		if c.Tier != TierC {
			continue
		}
		if _, err := os.Stat(c.GoldenPath); err == nil {
			t.Errorf("%s: must not have a golden (tier_c)", c.ID)
		}
	}
}

func TestCommittedScriptedGolden(t *testing.T) {
	root := repoCasesDir(t)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}

	var c *Case
	for i := range all {
		if all[i].ID == "default/tier_a_simple_set_a1" {
			c = &all[i]
			break
		}
	}
	if c == nil || !c.RunScript() {
		t.Fatal("missing scripted case default/tier_a_simple_set_a1")
	}

	for _, pass := range OpenSavePass(all) {
		if pass.ID == c.ID {
			t.Fatal("OpenSavePass must still skip the Office.js case")
		}
	}

	st, err := os.Stat(c.GoldenPath)
	if err != nil {
		t.Fatalf("committed scripted golden missing: %v", err)
	}
	if st.Size() == 0 {
		t.Fatal("committed scripted golden is empty")
	}
	m, err := golden.Read(c.GoldenPath)
	if err != nil {
		t.Fatal(err)
	}
	if m.Host != excel.HostID {
		t.Fatalf("host %q, want %s", m.Host, excel.HostID)
	}
	if m.Script != ScriptFile {
		t.Fatalf("script %q, want %s", m.Script, ScriptFile)
	}
	if m.Input != InitFile {
		t.Fatalf("input %q, want %s", m.Input, InitFile)
	}
	if m.ExcelVersion == "" || m.ExcelBuild == "" {
		t.Fatal("missing Excel version/build in scripted sidecar")
	}
	if m.GeneratedAt == "" {
		t.Fatal("missing generatedAt")
	}

	if !xlsxContains(t, c.GoldenPath, "calipers") {
		t.Fatal("scripted golden.xlsx must contain the Office.js value calipers")
	}
	if xlsxContains(t, c.InitPath, "calipers") {
		t.Fatal("tier_a_simple_set_a1 init.xlsx must not already contain calipers")
	}
}

func TestScratchSuite(t *testing.T) {
	root := repoCasesDir(t)
	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	var foundScratch bool
	for _, s := range corpus.Suites {
		if s == "scratch" {
			foundScratch = true
			break
		}
	}
	if !foundScratch {
		t.Fatal("corpus is missing suite scratch")
	}

	scratch, err := FilterSuite(corpus.Cases, "scratch", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if n := len(scratch); n < 10 || n > 20 {
		t.Fatalf("scratch cases: %d, want 10-20", n)
	}
	for _, c := range scratch {
		if c.Suite != "scratch" || !strings.HasPrefix(c.ID, "scratch/") {
			t.Errorf("FilterSuite scratch returned %s (suite=%q)", c.ID, c.Suite)
		}
		if strings.HasPrefix(c.Name, "tier_") || c.Tier != "" {
			t.Errorf("%s: scratch case names must not use a tier_ prefix (tier=%q)", c.ID, c.Tier)
		}
	}

	emptyInit, err := os.ReadFile(filepath.Join(root, "roundtrip", "tier_a_empty", InitFile))
	if err != nil {
		t.Fatal(err)
	}

	var bodies []byte
	for _, c := range scratch {
		if !c.RunScript() {
			t.Errorf("%s: scratch case must have a non-empty script.js", c.ID)
			continue
		}
		initData, err := os.ReadFile(c.InitPath)
		if err != nil {
			t.Errorf("%s: init: %v", c.ID, err)
			continue
		}
		if !bytes.Equal(initData, emptyInit) {
			t.Errorf("%s: init.xlsx must be a byte-identical copy of roundtrip/tier_a_empty", c.ID)
		}
		st, err := os.Stat(c.GoldenPath)
		if err != nil || st.Size() == 0 {
			t.Errorf("%s: scratch must have a committed golden.xlsx", c.ID)
		}
		body, err := os.ReadFile(c.ScriptPath)
		if err != nil {
			t.Errorf("%s: script: %v", c.ID, err)
			continue
		}
		if len(bytes.TrimSpace(body)) == 0 {
			t.Errorf("%s: script.js is empty; discovery would skip it", c.ID)
			continue
		}
		if !bytes.Contains(body, []byte("Excel.run")) {
			t.Errorf("%s: script.js is not Office.js Excel.run:\n%s", c.ID, body)
		}
		bodies = append(bodies, body...)
		bodies = append(bodies, '\n')
	}

	joined := string(bodies)
	for _, needle := range []string{
		"pivotTables.add",
		"tables.add",
		"charts.add",
		"conditionalFormats",
		"columnWidth",
		"rowHeight",
	} {
		if !strings.Contains(joined, needle) {
			t.Errorf("scratch suite missing %s", needle)
		}
	}
	spill := regexp.MustCompile(`SEQUENCE|FILTER|UNIQUE|SORT|RANDARRAY`)
	if !spill.MatchString(joined) {
		t.Error("scratch suite missing a spilling dynamic-array formula (SEQUENCE/FILTER/UNIQUE/SORT/RANDARRAY)")
	}
	if !strings.Contains(joined, "freezePanes") && !strings.Contains(joined, "showGridlines") && !strings.Contains(joined, "tab.color") {
		t.Error("scratch suite missing a worksheet/workbook setting (freezePanes, showGridlines, or tab.color)")
	}
	if !regexp.MustCompile(`\.values\s*=\s*\[[\s\S]*?"[A-Za-z]`).Match(bodies) {
		t.Error("scratch suite missing text cell values")
	}
	if !regexp.MustCompile(`\.values\s*=\s*\[[\s\S]*?[0-9]`).Match(bodies) {
		t.Error("scratch suite missing numeric cell values")
	}

	for _, c := range OpenSavePass(corpus.Cases) {
		if c.Suite == "scratch" {
			t.Errorf("OpenSavePass included scratch %s", c.ID)
		}
	}
	golden := GoldenComparePass(corpus.Cases)
	var nScratchGolden int
	for _, c := range golden {
		if c.Suite == "scratch" {
			nScratchGolden++
		}
	}
	if nScratchGolden != len(scratch) {
		t.Errorf("GoldenComparePass scratch count=%d, want %d", nScratchGolden, len(scratch))
	}
	def := DefaultPass(corpus.Cases)
	var nScratchDefault int
	for _, c := range def {
		if c.Suite == "scratch" {
			nScratchDefault++
		}
	}
	if nScratchDefault != len(scratch) {
		t.Fatalf("DefaultPass scratch count=%d, want %d (unprefixed scratch included)", nScratchDefault, len(scratch))
	}
}

func TestLoadUnprefixedCase(t *testing.T) {
	root := t.TempDir()
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeCase(t, filepath.Join(root, "scratch"), "text", "pk", &js)
	if err := os.Mkdir(filepath.Join(root, "scratch", "not_a_case"), 0o755); err != nil {
		t.Fatal(err)
	}

	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "scratch/text" || got[0].Name != "text" || got[0].Tier != "" || got[0].Suite != "scratch" {
		t.Fatalf("got %+v", got)
	}
	if !got[0].RunScript() {
		t.Fatal("unprefixed scratch case should run its script")
	}
	pass := DefaultPass(got)
	if len(pass) != 1 || pass[0].ID != "scratch/text" {
		t.Fatalf("DefaultPass should include unprefixed scratch, got %+v", pass)
	}
}

func TestLoadFlatUnprefixedCases(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async () => {});\n"
	writeCase(t, root, "text", "pk", &js)
	writeCase(t, root, "table", "pk", &js)
	if err := os.Mkdir(filepath.Join(root, "not_a_case"), 0o755); err != nil {
		t.Fatal(err)
	}

	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(corpus.Suites) != 0 {
		t.Fatalf("flat unprefixed layout should not report suites, got %v", corpus.Suites)
	}
	if len(corpus.Cases) != 2 {
		t.Fatalf("len=%d, want 2", len(corpus.Cases))
	}
	byID := map[string]Case{}
	for _, c := range corpus.Cases {
		byID[c.ID] = c
		if c.Suite != "" || c.Tier != "" {
			t.Errorf("%s: suite=%q tier=%q, want empty", c.ID, c.Suite, c.Tier)
		}
	}
	if _, ok := byID["text"]; !ok {
		t.Fatalf("missing text: %+v", byID)
	}
	if _, ok := byID["table"]; !ok {
		t.Fatalf("missing table: %+v", byID)
	}
}

func xlsxContains(t *testing.T, path, needle string) bool {
	t.Helper()
	r, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("zip %s: %v", path, err)
	}
	defer r.Close()
	want := []byte(needle)
	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		if bytes.Contains(data, want) {
			return true
		}
	}
	return false
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
