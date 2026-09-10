package scripts_test

import (
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
)

func scriptsDir(t *testing.T) string {
	t.Helper()
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("runtime.Caller failed")
	}
	return filepath.Dir(file)
}

func writeExec(t *testing.T, path, body string) {
	t.Helper()
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(body), 0o755); err != nil {
		t.Fatal(err)
	}
}

func readFile(t *testing.T, path string) string {
	t.Helper()
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return string(b)
}

func engineFromLog(t *testing.T, log string) string {
	t.Helper()
	for _, line := range strings.Split(log, "\n") {
		if !strings.Contains(line, "verify") || !strings.Contains(line, "--engine") {
			continue
		}
		fields := strings.Fields(line)
		for i, f := range fields {
			if f == "--engine" && i+1 < len(fields) {
				return fields[i+1]
			}
		}
	}
	t.Fatalf("no calipers verify --engine in log:\n%s", log)
	return ""
}

func gitStub(logPath string) string {
	return `#!/bin/sh
echo "git $*" >> "` + logPath + `"
if [ "$1" = "clone" ]; then
  dest="$3"
  mkdir -p "$dest/.git"
  echo "ref: refs/heads/main" > "$dest/.git/HEAD"
fi
exit 0
`
}

func cargoStub(logPath string) string {
	return `#!/bin/sh
echo "cargo $*" >> "` + logPath + `"
mkdir -p target-native/debug
cat > target-native/debug/mog <<'EOF'
#!/bin/sh
if [ "$1" = "--help" ] || [ "$1" = "-h" ] || [ -z "$1" ]; then
  printf '%s\n' 'Usage: mog <script.js>' '       mog --eval <source>'
  exit 0
fi
exit 0
EOF
chmod +x target-native/debug/mog
exit 0
`
}

func bazelStub(logPath string) string {
	return `#!/bin/sh
echo "bazel $*" >> "` + logPath + `"
mkdir -p dist/cli
cat > dist/cli/cells <<EOF
#!/bin/sh
echo "cells \$*" >> "` + logPath + `"
in=""
out=""
while [ \$# -gt 0 ]; do
  case "\$1" in
    -i) in="\$2"; shift 2 ;;
    -y|--eval|-q) shift ;;
    --script) shift 2 ;;
    *)
      if [ -z "\$out" ]; then out="\$1"; fi
      shift
      ;;
  esac
done
if [ -n "\$in" ] && [ -n "\$out" ]; then
  cp "\$in" "\$out"
fi
exit 0
EOF
chmod +x dist/cli/cells
exit 0
`
}

func calipersStub(logPath string) string {
	return `#!/bin/sh
echo "calipers $*" >> "` + logPath + `"
echo "calculation policy: host default (no recalculation requested)"
echo "[1/1] roundtrip/simple PASS (package match)"
echo "verify: 1 pass, 0 fail, 0 error"
exit 0
`
}

func runScript(t *testing.T, script, stubDir, calipersBin string, extraEnv []string, args ...string) []byte {
	t.Helper()
	cmd := exec.Command("bash", append([]string{script}, args...)...)
	cmd.Dir = filepath.Dir(scriptsDir(t))
	env := append(os.Environ(),
		"PATH="+stubDir+string(os.PathListSeparator)+os.Getenv("PATH"),
		"CALIPERS_BIN="+calipersBin,
	)
	env = append(env, extraEnv...)
	cmd.Env = env
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("%s: %v\n%s", filepath.Base(script), err, out)
	}
	return out
}

func TestVerifyMogScriptClonesBuildsAndVerifies(t *testing.T) {
	if runtime.GOOS == "windows" {
		t.Skip("verify-mog.sh stubs are POSIX executables")
	}
	dir := t.TempDir()
	stubDir := filepath.Join(dir, "bin")
	logPath := filepath.Join(dir, "stub.log")
	cloneDir := filepath.Join(dir, "mog")
	calipers := filepath.Join(stubDir, "calipers")
	writeExec(t, filepath.Join(stubDir, "git"), gitStub(logPath))
	writeExec(t, filepath.Join(stubDir, "cargo"), cargoStub(logPath))
	writeExec(t, calipers, calipersStub(logPath))

	script := filepath.Join(scriptsDir(t), "verify-mog.sh")
	env := []string{"MOG_CLONE_DIR=" + cloneDir, "STUB_LOG=" + logPath}

	out1 := runScript(t, script, stubDir, calipers, env, "--case", "roundtrip/simple")
	out2 := runScript(t, script, stubDir, calipers, env, "--case", "roundtrip/simple")
	log := readFile(t, logPath)

	if !strings.Contains(log, "github.com/fundamental-research-labs/mog") {
		t.Fatalf("clone of mog missing:\n%s", log)
	}
	if strings.Count(log, "git clone") != 1 {
		t.Fatalf("expected one git clone (second run should reuse .git):\n%s", log)
	}
	if !strings.Contains(log, "cargo") || !strings.Contains(log, "build") || !strings.Contains(log, "-p mog") {
		t.Fatalf("cargo build -p mog missing:\n%s", log)
	}
	if !strings.Contains(log, "verify") || !strings.Contains(log, "--engine") {
		t.Fatalf("calipers verify --engine missing:\n%s", log)
	}
	if !strings.Contains(log, "--case") || !strings.Contains(log, "roundtrip/simple") {
		t.Fatalf("extra args not forwarded to verify:\n%s", log)
	}
	engine := engineFromLog(t, log)
	if !strings.HasPrefix(engine, cloneDir) {
		t.Fatalf("engine path %q is not from the mog build at %s\nlog:\n%s", engine, cloneDir, log)
	}
	if _, err := os.Stat(engine); err != nil {
		t.Fatalf("engine %s: %v", engine, err)
	}
	combined := string(out1) + string(out2)
	if !strings.Contains(combined, "PASS") && !strings.Contains(log, "verify") {
		t.Fatalf("verify did not report outcomes:\n%s\n%s", combined, log)
	}

	in := filepath.Join(dir, "in.xlsx")
	outX := filepath.Join(dir, "out.xlsx")
	scriptJS := filepath.Join(dir, "script.js")
	if err := os.WriteFile(in, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(scriptJS, []byte("await Excel.run(async () => {});"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := exec.Command(engine, "save", in, outX).Run(); err != nil {
		t.Fatalf("engine save: %v", err)
	}
	if _, err := os.Stat(outX); err != nil {
		t.Fatalf("save did not write %s: %v", outX, err)
	}
	outRun := filepath.Join(dir, "run.xlsx")
	if err := exec.Command(engine, "run", in, scriptJS, outRun).Run(); err != nil {
		t.Fatalf("engine run: %v", err)
	}
	if _, err := os.Stat(outRun); err != nil {
		t.Fatalf("run did not write %s: %v", outRun, err)
	}
}

func TestVerifyCellsScriptClonesBuildsAndVerifies(t *testing.T) {
	if runtime.GOOS == "windows" {
		t.Skip("verify-cells.sh stubs are POSIX executables")
	}
	dir := t.TempDir()
	stubDir := filepath.Join(dir, "bin")
	logPath := filepath.Join(dir, "stub.log")
	cloneDir := filepath.Join(dir, "cells")
	calipers := filepath.Join(stubDir, "calipers")
	writeExec(t, filepath.Join(stubDir, "git"), gitStub(logPath))
	writeExec(t, filepath.Join(stubDir, "bazel"), bazelStub(logPath))
	writeExec(t, calipers, calipersStub(logPath))

	script := filepath.Join(scriptsDir(t), "verify-cells.sh")
	env := []string{"CELLS_CLONE_DIR=" + cloneDir, "STUB_LOG=" + logPath}

	runScript(t, script, stubDir, calipers, env, "--case", "roundtrip/simple")
	runScript(t, script, stubDir, calipers, env, "--case", "roundtrip/simple")
	log := readFile(t, logPath)

	if !strings.Contains(log, "github.com/aduermael/cells") {
		t.Fatalf("clone of cells missing:\n%s", log)
	}
	if strings.Count(log, "git clone") != 1 {
		t.Fatalf("expected one git clone (second run should reuse .git):\n%s", log)
	}
	if !strings.Contains(log, "bazel") || !strings.Contains(log, "cli-headless-no-collab") {
		t.Fatalf("cells bazel build missing:\n%s", log)
	}
	if !strings.Contains(log, "verify") || !strings.Contains(log, "--engine") {
		t.Fatalf("calipers verify --engine missing:\n%s", log)
	}
	engine := engineFromLog(t, log)
	if !strings.HasPrefix(engine, cloneDir) {
		t.Fatalf("engine path %q is not from the cells build at %s\nlog:\n%s", engine, cloneDir, log)
	}

	in := filepath.Join(dir, "in.xlsx")
	outX := filepath.Join(dir, "out.xlsx")
	scriptJS := filepath.Join(dir, "script.js")
	if err := os.WriteFile(in, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(scriptJS, []byte("await Excel.run(async () => {});"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := exec.Command(engine, "save", in, outX).Run(); err != nil {
		t.Fatalf("engine save: %v", err)
	}
	if _, err := os.Stat(outX); err != nil {
		t.Fatalf("save did not write %s: %v", outX, err)
	}
	if !strings.Contains(readFile(t, logPath), "-i") {
		t.Fatalf("cells adapter save must invoke cells -i:\n%s", readFile(t, logPath))
	}
	outRun := filepath.Join(dir, "run.xlsx")
	if err := exec.Command(engine, "run", in, scriptJS, outRun).Run(); err != nil {
		t.Fatalf("engine run: %v", err)
	}
	if _, err := os.Stat(outRun); err != nil {
		t.Fatalf("run did not write %s: %v", outRun, err)
	}
	if !strings.Contains(readFile(t, logPath), "--script") {
		t.Fatalf("cells adapter run must invoke cells --script:\n%s", readFile(t, logPath))
	}
}
