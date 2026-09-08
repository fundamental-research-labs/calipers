package mog

import (
	"context"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

func TestRunScriptRequiresPath(t *testing.T) {
	h := NewHost()
	if err := h.RunScript("in.xlsx", "", "out.xlsx"); err == nil {
		t.Fatal("empty script should error")
	}
}

func TestOpenSaveRequiresXlsx(t *testing.T) {
	dir := t.TempDir()
	in := filepath.Join(dir, "in.xlsx")
	if err := os.WriteFile(in, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	h := NewHost()
	h.Bin = "mog"
	if err := h.OpenSave(in, filepath.Join(dir, "out.xls")); err == nil || !strings.Contains(err.Error(), ".xlsx") {
		t.Fatalf("error = %v, want .xlsx", err)
	}
}

func TestOpenSaveInvokesCLIWithoutScript(t *testing.T) {
	dir := t.TempDir()
	in := filepath.Join(dir, "in.xlsx")
	out := filepath.Join(dir, "out.xlsx")
	if err := os.WriteFile(in, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	var got []string
	old := execCommandContext
	execCommandContext = func(ctx context.Context, name string, args ...string) *exec.Cmd {
		got = append([]string{name}, args...)
		return exec.CommandContext(ctx, "cp", in, out)
	}
	defer func() { execCommandContext = old }()

	h := NewHost()
	h.Bin = "mog"
	if err := h.OpenSave(in, out); err != nil {
		t.Fatal(err)
	}
	if len(got) != 5 || got[1] != "--input" || got[3] != "--output" {
		t.Fatalf("args = %v, want mog --input in --output out", got)
	}
	if filepath.Base(got[4]) != "out.xlsx" {
		t.Fatalf("export path = %v", got)
	}
}

func TestRunScriptPassesScriptAfterExportFlags(t *testing.T) {
	dir := t.TempDir()
	in := filepath.Join(dir, "in.xlsx")
	out := filepath.Join(dir, "out.xlsx")
	script := filepath.Join(dir, "script.js")
	if err := os.WriteFile(in, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(script, []byte("await Excel.run(async () => {});"), 0o644); err != nil {
		t.Fatal(err)
	}
	var got []string
	old := execCommandContext
	execCommandContext = func(ctx context.Context, name string, args ...string) *exec.Cmd {
		got = args
		return exec.CommandContext(ctx, "cp", in, out)
	}
	defer func() { execCommandContext = old }()

	h := NewHost()
	h.Bin = "mog"
	if err := h.RunScript(in, script, out); err != nil {
		t.Fatal(err)
	}
	if len(got) != 5 || got[0] != "--input" || got[2] != "--output" {
		t.Fatalf("args = %v", got)
	}
	if filepath.Base(got[3]) != "out.xlsx" {
		t.Fatalf("export = %s", got[3])
	}
	if filepath.Base(got[4]) != "script.js" {
		t.Fatalf("script not passed: %v", got)
	}
}

func TestIsMogOSS(t *testing.T) {
	dir := t.TempDir()
	if isMogOSS(dir) {
		t.Fatal("empty dir is not OSS mog")
	}
	office := filepath.Join(dir, "compute", "officejs")
	if err := os.MkdirAll(office, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(office, "Cargo.toml"), []byte("[package]\nname = \"mog\"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if !isMogOSS(dir) {
		t.Fatal("want OSS mog crate at compute/officejs")
	}
}
