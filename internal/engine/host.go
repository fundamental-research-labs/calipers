// Package engine drives a caller-supplied workbook engine binary.
//
// The binary is provided at verify time. Calipers does not embed a particular
// engine. Argv contract (issue #6):
//
//	<engine> save <in.xlsx> <out.xlsx>
//	<engine> run  <in.xlsx> <script.js> <out.xlsx>
package engine

import (
	"context"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"time"
)

// DefaultTimeout is the max time for one engine save/run.
const DefaultTimeout = 2 * time.Minute

// Host execs an external engine binary with the same OpenSave/RunScript
// operations as the Excel host.
type Host struct {
	Path    string
	Timeout time.Duration
}

// New returns a host for the given executable path.
func New(path string) *Host {
	return &Host{Path: path}
}

func (h *Host) timeout() time.Duration {
	if h.Timeout > 0 {
		return h.Timeout
	}
	return DefaultTimeout
}

// OpenSave runs: engine save <input> <output>
func (h *Host) OpenSave(inputPath, outputPath string) error {
	in, out, err := prepareIO(inputPath, outputPath)
	if err != nil {
		return err
	}
	if err := h.exec("save", in, out); err != nil {
		return err
	}
	return wrote(out)
}

// RunScript runs: engine run <input> <script> <output>
func (h *Host) RunScript(inputPath, scriptPath, outputPath string) error {
	if strings.TrimSpace(scriptPath) == "" {
		return fmt.Errorf("script path is required")
	}
	in, out, err := prepareIO(inputPath, outputPath)
	if err != nil {
		return err
	}
	script, err := filepath.Abs(scriptPath)
	if err != nil {
		return fmt.Errorf("script path: %w", err)
	}
	if _, err := os.Stat(script); err != nil {
		return fmt.Errorf("script: %w", err)
	}
	if err := h.exec("run", in, script, out); err != nil {
		return err
	}
	return wrote(out)
}

func wrote(out string) error {
	if _, err := os.Stat(out); err != nil {
		return fmt.Errorf("engine did not write %s", out)
	}
	return nil
}

func (h *Host) exec(args ...string) error {
	if strings.TrimSpace(h.Path) == "" {
		return fmt.Errorf("engine binary path is required")
	}
	ctx, cancel := context.WithTimeout(context.Background(), h.timeout())
	defer cancel()
	cmd := execCommandContext(ctx, h.Path, args...)
	out, err := cmd.CombinedOutput()
	if err != nil {
		msg := strings.TrimSpace(string(out))
		if msg == "" {
			msg = err.Error()
		}
		return fmt.Errorf("engine: %s", msg)
	}
	return nil
}

var execCommandContext = exec.CommandContext

func prepareIO(inputPath, outputPath string) (string, string, error) {
	if strings.TrimSpace(inputPath) == "" || strings.TrimSpace(outputPath) == "" {
		return "", "", fmt.Errorf("input and output paths are required")
	}
	in, err := filepath.Abs(inputPath)
	if err != nil {
		return "", "", fmt.Errorf("input path: %w", err)
	}
	out, err := filepath.Abs(outputPath)
	if err != nil {
		return "", "", fmt.Errorf("output path: %w", err)
	}
	if !strings.EqualFold(filepath.Ext(out), ".xlsx") {
		return "", "", fmt.Errorf("output must be .xlsx, got %q", out)
	}
	if _, err := os.Stat(in); err != nil {
		return "", "", fmt.Errorf("input: %w", err)
	}
	if err := os.MkdirAll(filepath.Dir(out), 0o755); err != nil {
		return "", "", err
	}
	return in, out, nil
}
