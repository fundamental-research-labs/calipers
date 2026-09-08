package excel

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestStampWebExtension(t *testing.T) {
	src := corpusInit(t)
	dst := filepath.Join(t.TempDir(), "stamped.xlsx")
	if err := StampWebExtension(dst, src, AddinID); err != nil {
		t.Fatal(err)
	}
	r, err := zip.OpenReader(dst)
	if err != nil {
		t.Fatal(err)
	}
	defer r.Close()
	found := map[string][]byte{}
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
		found[f.Name] = data
	}
	for _, name := range []string{webextName, taskpanesName, taskpanesRels, contentTypesName} {
		if _, ok := found[name]; !ok {
			t.Fatalf("missing zip part %s", name)
		}
	}
	we := string(found[webextName])
	if !strings.Contains(we, AddinID) {
		t.Fatalf("webextension missing add-in id:\n%s", we)
	}
	if !strings.Contains(we, "Office.AutoShowTaskpaneWithDocument") {
		t.Fatal("webextension must auto-show the task pane")
	}
	ct := string(found[contentTypesName])
	if !strings.Contains(ct, "webextension+xml") {
		t.Fatalf("content types:\n%s", ct)
	}
}

func corpusInit(t *testing.T) string {
	t.Helper()
	p := filepath.Join(filepath.Dir(corpusScript(t)), "init.xlsx")
	if _, err := os.Stat(p); err != nil {
		t.Fatal(err)
	}
	return p
}
