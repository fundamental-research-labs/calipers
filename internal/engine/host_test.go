package engine

import (
	"context"
	"os"
	"os/exec"
	"path/filepath"
	"reflect"
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

func TestOpenSaveDoesNotAcceptStaleOutput(t *testing.T) {
	dir := t.TempDir()
	in := filepath.Join(dir, "in.xlsx")
	out := filepath.Join(dir, "out.xlsx")
	if err := os.WriteFile(in, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(out, []byte("stale-export"), 0o644); err != nil {
		t.Fatal(err)
	}
	old := execCommandContext
	execCommandContext = func(ctx context.Context, name string, args ...string) *exec.Cmd {
		// Exit 0 without creating or replacing the output.
		return exec.CommandContext(ctx, "true")
	}
	defer func() { execCommandContext = old }()

	h := New("/bin/engine")
	err := h.OpenSave(in, out)
	if err == nil {
		t.Fatal("OpenSave succeeded using leftover output; engine did not write the file")
	}
	if !strings.Contains(err.Error(), "did not write") {
		t.Fatalf("error = %v, want did not write", err)
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

func TestRecalculateArgv(t *testing.T) {
	for _, command := range []string{"save", "run"} {
		t.Run(command, func(t *testing.T) {
			dir := t.TempDir()
			in, out, script := filepath.Join(dir, "in.xlsx"), filepath.Join(dir, "out.xlsx"), filepath.Join(dir, "script.js")
			for _, path := range []string{in, script} {
				if err := os.WriteFile(path, []byte("fixture"), 0o644); err != nil {
					t.Fatal(err)
				}
			}
			var got []string
			old := execCommandContext
			execCommandContext = func(ctx context.Context, name string, args ...string) *exec.Cmd {
				got = append([]string{name}, args...)
				return exec.CommandContext(ctx, "cp", in, out)
			}
			t.Cleanup(func() { execCommandContext = old })
			h := New("/bin/engine")
			h.Recalculate = true
			want := []string{h.Path, command, "--recalculate", in}
			var err error
			if command == "save" {
				err = h.OpenSave(in, out)
			} else {
				want = append(want, script)
				err = h.RunScript(in, script, out)
			}
			want = append(want, out)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(got, want) {
				t.Fatalf("argv = %v, want %v", got, want)
			}
		})
	}
}
