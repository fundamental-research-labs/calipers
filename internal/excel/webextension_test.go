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
	if !strings.Contains(we, `store="developer"`) || !strings.Contains(we, `storeType="Registry"`) {
		t.Fatalf("webextension must resolve via WEF\\Developer, got:\n%s", we)
	}
	if _, ok := found["xl/webextensions/webextension1.xml"]; ok {
		t.Fatal("stamp must use official sideload part webextension.xml, not webextension1.xml")
	}
	reg, err := NewSideloadReg(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	if reg.DeveloperName != AddinID || !strings.Contains(reg.DeveloperKey, `WEF\Developer`) {
		t.Fatalf("Developer registration does not match stamp: %+v", reg)
	}
	ct := string(found[contentTypesName])
	if !strings.Contains(ct, "webextension+xml") {
		t.Fatalf("content types:\n%s", ct)
	}
	pkgRels := string(found[packageRelsName])
	if !strings.Contains(pkgRels, "webextensiontaskpanes") || !strings.Contains(pkgRels, "/xl/webextensions/taskpanes.xml") {
		t.Fatalf("package rels must point at taskpanes (Office sideload template):\n%s", pkgRels)
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
