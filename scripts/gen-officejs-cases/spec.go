package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

type spec struct {
	name   string
	script string
}

func sheetJS(name, comment, inner string) spec {
	inner = strings.TrimRight(inner, "\n")
	return spec{
		name:   name,
		script: fmt.Sprintf("// %s\nawait Excel.run(async (context) => {\n  const sheet = context.workbook.worksheets.getActiveWorksheet();\n%s\n  await context.sync();\n});\n", comment, inner),
	}
}

func rawJS(name, script string) spec {
	if !strings.HasSuffix(script, "\n") {
		script += "\n"
	}
	return spec{name: name, script: script}
}

func allSpecs() []spec {
	out := make([]spec, 0, 100)
	out = append(out, rangeSpecs()...)
	out = append(out, styleSpecs()...)
	out = append(out, tableSpecs()...)
	out = append(out, chartSpecs()...)
	out = append(out, spillSpecs()...)
	out = append(out, restSpecs()...)
	return out
}

func generate(emptyInit, destDir string) (int, error) {
	data, err := os.ReadFile(emptyInit)
	if err != nil {
		return 0, fmt.Errorf("read empty init: %w", err)
	}
	specs := allSpecs()
	if err := os.MkdirAll(destDir, 0o755); err != nil {
		return 0, err
	}
	for _, s := range specs {
		if s.name == "" || strings.TrimSpace(s.script) == "" {
			return 0, fmt.Errorf("invalid spec %q", s.name)
		}
		dir := filepath.Join(destDir, s.name)
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return 0, err
		}
		if err := os.WriteFile(filepath.Join(dir, "init.xlsx"), data, 0o644); err != nil {
			return 0, err
		}
		if err := os.WriteFile(filepath.Join(dir, "script.js"), []byte(s.script), 0o644); err != nil {
			return 0, err
		}
	}
	return len(specs), nil
}
