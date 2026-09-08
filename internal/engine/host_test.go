package engine

import (
	"context"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

func TestRunScriptRequiresPath(t *testing.T) {
	h := New("engine")
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
	h := New("engine")
	if err := h.OpenSave(in, filepath.Join(dir, "out.xls")); err == nil || !strings.Contains(err.Error(), ".xlsx") {
		t.Fatalf("error = %v, want .xlsx", err)
	}
}

func TestOpenSaveArgv(t *testing.T) {
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

	h := New("/bin/engine")
	if err := h.OpenSave(in, out); err != nil {
		t.Fatal(err)
	}
	if len(got) != 4 || got[0] != "/bin/engine" || got[1] != "save" {
		t.Fatalf("argv = %v, want engine save in out", got)
	}
	if filepath.Base(got[2]) != "in.xlsx" || filepath.Base(got[3]) != "out.xlsx" {
		t.Fatalf("paths = %v", got)
	}
}

func TestRunScriptArgv(t *testing.T) {
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

	h := New("/bin/engine")
	if err := h.RunScript(in, script, out); err != nil {
		t.Fatal(err)
	}
	if len(got) != 4 || got[0] != "run" {
		t.Fatalf("argv = %v, want run in script out", got)
	}
	if filepath.Base(got[1]) != "in.xlsx" || filepath.Base(got[2]) != "script.js" || filepath.Base(got[3]) != "out.xlsx" {
		t.Fatalf("argv = %v", got)
	}
}
