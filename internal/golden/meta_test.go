package golden

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"
)

func TestWriteRead(t *testing.T) {
	dir := t.TempDir()
	xlsx := filepath.Join(dir, "golden.xlsx")
	if err := os.WriteFile(xlsx, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	m := New("excel-win", "windows", "16.0", "17628", "simple.xlsx")
	if err := Write(xlsx, m); err != nil {
		t.Fatal(err)
	}
	got, err := Read(xlsx)
	if err != nil {
		t.Fatal(err)
	}
	if got.Host != "excel-win" || got.ExcelVersion != "16.0" || got.ExcelBuild != "17628" {
		t.Fatalf("got %+v", got)
	}
	if got.Input != "simple.xlsx" || got.Tool != "calipers" || got.OS != "windows" {
		t.Fatalf("got %+v", got)
	}
	if got.GeneratedAt == "" {
		t.Fatal("missing generatedAt")
	}
	if got.Script != "" {
		t.Fatalf("sidecar without script invented Script=%q", got.Script)
	}
	raw, err := os.ReadFile(PathFor(xlsx))
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Contains(raw, []byte(`"script"`)) {
		t.Fatalf("load+save sidecar must omit script key, got:\n%s", raw)
	}
}

func TestWriteReadScriptIdentity(t *testing.T) {
	dir := t.TempDir()
	xlsx := filepath.Join(dir, "golden.xlsx")
	if err := os.WriteFile(xlsx, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	m := New("excel-win", "windows", "16.0", "17628", "init.xlsx")
	m.Script = "script.js"
	if err := Write(xlsx, m); err != nil {
		t.Fatal(err)
	}
	got, err := Read(xlsx)
	if err != nil {
		t.Fatal(err)
	}
	if got.Script != "script.js" {
		t.Fatalf("script identity = %q, want script.js", got.Script)
	}
	if got.Input != "init.xlsx" {
		t.Fatalf("input = %q", got.Input)
	}
	raw, err := os.ReadFile(PathFor(xlsx))
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(raw, []byte(`"script": "script.js"`)) {
		t.Fatalf("expected script identity in sidecar, got:\n%s", raw)
	}
}
