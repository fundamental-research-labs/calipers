// Package mog runs the open-source Mog CLI (cargo build -p mog).
package mog

import (
	"context"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"time"
)

// DefaultTimeout is the max time for one mog load/script/export.
const DefaultTimeout = 2 * time.Minute

const buildTimeout = 15 * time.Minute

// Host loads an xlsx in the Mog CLI, optionally runs Office.js, and exports.
//
//	mog --input <init.xlsx> --output <result.xlsx>
//	mog --input <init.xlsx> --output <result.xlsx> <script.js>
type Host struct {
	// Bin is the mog executable. Empty means discover (MOG_BIN, PATH) or
	// cargo-build from MOG_ROOT / Host.Root.
	Bin string
	// Root is the open-source mog checkout used for `cargo build -p mog`.
	Root    string
	Timeout time.Duration

	resolved string
}

// NewHost returns a Mog CLI host.
func NewHost() *Host {
	return &Host{}
}

func (h *Host) timeout() time.Duration {
	if h.Timeout > 0 {
		return h.Timeout
	}
	return DefaultTimeout
}

// OpenSave loads inputPath in mog and exports a new xlsx to outputPath.
func (h *Host) OpenSave(inputPath, outputPath string) error {
	return h.invoke(inputPath, outputPath, "")
}

// RunScript loads inputPath, runs a non-empty Office.js file, then exports.
func (h *Host) RunScript(inputPath, scriptPath, outputPath string) error {
	if strings.TrimSpace(scriptPath) == "" {
		return fmt.Errorf("script path is required")
	}
	return h.invoke(inputPath, outputPath, scriptPath)
}

func (h *Host) invoke(inputPath, outputPath, scriptPath string) error {
	absIn, err := filepath.Abs(inputPath)
	if err != nil {
		return fmt.Errorf("input path: %w", err)
	}
	absOut, err := filepath.Abs(outputPath)
	if err != nil {
		return fmt.Errorf("output path: %w", err)
	}
	if !strings.EqualFold(filepath.Ext(absOut), ".xlsx") {
		return fmt.Errorf("output must be .xlsx, got %q", absOut)
	}
	if _, err := os.Stat(absIn); err != nil {
		return fmt.Errorf("input: %w", err)
	}
	if err := os.MkdirAll(filepath.Dir(absOut), 0o755); err != nil {
		return err
	}

	args := []string{"--input", absIn, "--output", absOut}
	if scriptPath != "" {
		absScript, err := filepath.Abs(scriptPath)
		if err != nil {
			return fmt.Errorf("script path: %w", err)
		}
		if _, err := os.Stat(absScript); err != nil {
			return fmt.Errorf("script: %w", err)
		}
		args = append(args, absScript)
	}

	bin, err := h.resolveBin()
	if err != nil {
		return err
	}

	ctx, cancel := context.WithTimeout(context.Background(), h.timeout())
	defer cancel()
	cmd := execCommandContext(ctx, bin, args...)
	out, err := cmd.CombinedOutput()
	if err != nil {
		msg := strings.TrimSpace(string(out))
		if msg == "" {
			msg = err.Error()
		}
		return fmt.Errorf("mog: %s", msg)
	}
	if _, err := os.Stat(absOut); err != nil {
		return fmt.Errorf("mog did not write %s", absOut)
	}
	return nil
}

var execCommandContext = exec.CommandContext
var lookPath = exec.LookPath

func (h *Host) resolveBin() (string, error) {
	if h.resolved != "" {
		return h.resolved, nil
	}
	if h.Bin != "" {
		h.resolved = h.Bin
		return h.Bin, nil
	}
	if b := os.Getenv("MOG_BIN"); b != "" {
		h.resolved = b
		return b, nil
	}
	if p, err := lookPath("mog"); err == nil {
		h.resolved = p
		return p, nil
	}
	bin, err := h.buildCLI()
	if err != nil {
		return "", err
	}
	h.resolved = bin
	return bin, nil
}

func (h *Host) buildCLI() (string, error) {
	root, err := h.mogRoot()
	if err != nil {
		return "", err
	}
	ctx, cancel := context.WithTimeout(context.Background(), buildTimeout)
	defer cancel()
	cmd := execCommandContext(ctx, "cargo", "build", "-p", "mog")
	cmd.Dir = root
	out, err := cmd.CombinedOutput()
	if err != nil {
		msg := strings.TrimSpace(string(out))
		if msg == "" {
			msg = err.Error()
		}
		return "", fmt.Errorf("cargo build -p mog: %s", msg)
	}
	bin, err := findBuiltMog(root)
	if err != nil {
		return "", err
	}
	return bin, nil
}

func (h *Host) mogRoot() (string, error) {
	if h.Root != "" {
		return h.Root, nil
	}
	if r := os.Getenv("MOG_ROOT"); r != "" {
		return r, nil
	}
	var starts []string
	if wd, err := os.Getwd(); err == nil {
		starts = append(starts, wd)
	}
	seen := map[string]bool{}
	for _, start := range starts {
		for dir := start; ; dir = filepath.Dir(dir) {
			if seen[dir] {
				break
			}
			seen[dir] = true
			if isMogOSS(dir) {
				return dir, nil
			}
			parent := filepath.Dir(dir)
			if parent == dir {
				break
			}
		}
	}
	return "", fmt.Errorf("mog CLI not found: set MOG_BIN or MOG_ROOT (open-source mog checkout; cargo build -p mog)")
}

func isMogOSS(dir string) bool {
	p := filepath.Join(dir, "compute", "officejs", "Cargo.toml")
	data, err := os.ReadFile(p)
	if err != nil {
		return false
	}
	return strings.Contains(string(data), `name = "mog"`)
}

func findBuiltMog(root string) (string, error) {
	name := "mog"
	if runtime.GOOS == "windows" {
		name = "mog.exe"
	}
	candidates := []string{
		filepath.Join(root, "target-native", "debug", name),
		filepath.Join(root, "target-native", "release", name),
		filepath.Join(root, "target", "debug", name),
		filepath.Join(root, "target", "release", name),
	}
	for _, pattern := range []string{
		filepath.Join(root, "target-native", "*", "debug", name),
		filepath.Join(root, "target-native", "*", "release", name),
	} {
		matches, _ := filepath.Glob(pattern)
		candidates = append(candidates, matches...)
	}
	for _, p := range candidates {
		if st, err := os.Stat(p); err == nil && !st.IsDir() {
			return p, nil
		}
	}
	return "", fmt.Errorf("built mog binary not found under %s (cargo build -p mog)", root)
}
