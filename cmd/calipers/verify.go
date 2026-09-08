package main

import (
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/compare"
	"github.com/fundamental-research-labs/calipers/internal/mog"
)

// engine is the mog load/export surface used by verify.
type engine interface {
	OpenSave(inputPath, outputPath string) error
	RunScript(inputPath, scriptPath, outputPath string) error
}

var newVerifyEngine = func() engine { return mog.NewHost() }

const verifyUsage = `calipers verify [--cases-dir DIR] [--out-dir DIR] [--case ID]...

  Load each case in mog, run Office.js when present and non-empty,
  export a result xlsx (not the golden), and semantically compare
  to the case golden.

  Default walk is tier_a and tier_b (skip tier_c hostiles).
`

func verifyCmd(args []string) error {
	fs := flag.NewFlagSet("verify", flag.ContinueOnError)
	fs.SetOutput(os.Stderr)
	casesDir := fs.String("cases-dir", cases.DirName, "verification cases directory")
	outDir := fs.String("out-dir", "", "directory for mog exports (never the case golden)")
	var caseIDs []string
	fs.Func("case", "case id to run (repeatable or comma-separated; default: tier_a and tier_b)", func(s string) error {
		for _, id := range strings.Split(s, ",") {
			id = strings.TrimSpace(id)
			if id != "" {
				caseIDs = append(caseIDs, id)
			}
		}
		return nil
	})
	fs.Usage = func() {
		fmt.Fprint(os.Stdout, verifyUsage)
	}
	if err := fs.Parse(args); err != nil {
		if err == flag.ErrHelp {
			return nil
		}
		return err
	}
	if fs.NArg() > 1 {
		return fmt.Errorf("usage: calipers verify [--cases-dir DIR] [--out-dir DIR] [--case ID]...")
	}
	if fs.NArg() == 1 {
		*casesDir = fs.Arg(0)
	}
	return runVerify(newVerifyEngine(), *casesDir, caseIDs, *outDir, os.Stdout)
}

type caseOutcome struct {
	ID     string
	Export string
	Status string // pass, fail, error
	Detail string
}

func runVerify(eng engine, casesDir string, caseIDs []string, outDir string, w io.Writer) error {
	all, err := cases.Load(casesDir)
	if err != nil {
		return err
	}
	selected, err := cases.Select(all, caseIDs)
	if err != nil {
		return err
	}
	if outDir == "" {
		outDir = filepath.Join(os.TempDir(), "calipers-verify")
	}
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		return err
	}

	if len(selected) == 0 {
		fmt.Fprintf(w, "verify: 0 pass, 0 fail, 0 error\n")
		return nil
	}

	var nPass, nFail, nErr int
	for i, c := range selected {
		o := verifyOne(eng, c, outDir)
		switch o.Status {
		case "pass":
			nPass++
		case "fail":
			nFail++
		default:
			nErr++
		}
		fmt.Fprintf(w, "[%d/%d] %s %s", i+1, len(selected), c.ID, strings.ToUpper(o.Status))
		if o.Detail != "" {
			fmt.Fprintf(w, " %s", o.Detail)
		}
		fmt.Fprintln(w)
	}
	fmt.Fprintf(w, "verify: %d pass, %d fail, %d error\n", nPass, nFail, nErr)
	if nFail+nErr > 0 {
		return fmt.Errorf("%d failed, %d error", nFail, nErr)
	}
	return nil
}

func verifyOne(eng engine, c cases.Case, outDir string) caseOutcome {
	exportPath := filepath.Join(outDir, c.ID+".xlsx")
	if filepath.Clean(exportPath) == filepath.Clean(c.GoldenPath) {
		return caseOutcome{ID: c.ID, Export: exportPath, Status: "error", Detail: "refusing to overwrite golden"}
	}

	var err error
	if c.RunScript() {
		err = eng.RunScript(c.InitPath, c.ScriptPath, exportPath)
	} else {
		err = eng.OpenSave(c.InitPath, exportPath)
	}
	if err != nil {
		return caseOutcome{ID: c.ID, Export: exportPath, Status: "error", Detail: err.Error()}
	}

	got, err := compare.Files(exportPath, c.GoldenPath)
	if err != nil {
		return caseOutcome{ID: c.ID, Export: exportPath, Status: "error", Detail: err.Error()}
	}
	if got.Equal {
		return caseOutcome{ID: c.ID, Export: exportPath, Status: "pass"}
	}
	detail := fmt.Sprintf("(%d diffs)", len(got.Diffs))
	if len(got.Diffs) > 0 {
		detail += " " + got.Diffs[0].Part
	}
	return caseOutcome{ID: c.ID, Export: exportPath, Status: "fail", Detail: detail}
}
