package main

import (
	"fmt"

	"github.com/fundamental-research-labs/calipers/internal/cases"
)

// executeCase is shared by verification, Excel golden generation, and measurement.
func executeCase(host engine, c cases.Case, output string) error {
	if c.Recalculate {
		if c.RunScript() {
			return fmt.Errorf("case %s: recalculate requires an unscripted case", c.ID)
		}
		h, ok := host.(interface{ RecalculateOpenSave(string, string) error })
		if !ok {
			return fmt.Errorf("case %s: host does not support explicit recalculation", c.ID)
		}
		return h.RecalculateOpenSave(c.InitPath, output)
	}
	if c.RunScript() {
		return host.RunScript(c.InitPath, c.ScriptPath, output)
	}
	return host.OpenSave(c.InitPath, output)
}
