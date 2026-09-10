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
	enginehost "github.com/fundamental-research-labs/calipers/internal/engine"
	"github.com/fundamental-research-labs/calipers/internal/excel"
	"github.com/fundamental-research-labs/calipers/internal/xlsxmodel"
)

// engine is the OpenSave/RunScript surface shared by Excel and an external binary.
type engine interface {
	OpenSave(inputPath, outputPath string) error
	RunScript(inputPath, scriptPath, outputPath string) error
}

var (
	newExcelHost  = func() engine { return excel.NewHost() }
	newBinaryHost = func(path string, recalculate bool) engine {
		h := enginehost.New(path)
		h.Recalculate = recalculate
		return h
	}
)

const verifyUsage = `calipers verify --engine <path|excel> [--recalculate] [--package] [--cases-dir DIR] [--out-dir DIR] [--suite NAME] [--case ID]...

  For each case: run the engine (load init.xlsx, Office.js if present and
  non-empty), export a result xlsx (not the golden), compare resolved
  workbook semantics to the committed Excel golden. PASS/FAIL is semantic
  diffs (values, types, formulas, styles, sheets, names, merges, freeze,
  date1904), not ZIP-part counts.

  --engine excel    Excel COM host (Windows)
  --engine PATH     external binary:  save <in> <out>
                                      run  <in> <script.js> <out>
  --recalculate     request full recalculation before export by passing
                    --recalculate after save/run; requires engine support.
                    Unsupported with --engine excel. Default: host policy
                    (no recalculation requested).
  --package         also print ZIP-package diffs (diagnostic only; does
                    not change PASS/FAIL)
  --suite NAME      run only this suite directory under --cases-dir
                    (e.g. roundtrip, default, scratch)
  --case ID         run only this case (repeatable or comma-separated;
                    id is suite/name, e.g. roundtrip/simple)

  --engine is required. Default walk is cases that have a committed
  golden.xlsx (skip cases with no golden).
`

func verifyCmd(args []string) error {
	fs := flag.NewFlagSet("verify", flag.ContinueOnError)
	fs.SetOutput(os.Stderr)
	engineSpec := fs.String("engine", "", "engine binary path, or 'excel'")
	recalculate := fs.Bool("recalculate", false, "request recalculation before export (external engines supporting --recalculate only)")
	packageDiag := fs.Bool("package", false, "print ZIP-package diffs without changing PASS/FAIL")
	casesDir := fs.String("cases-dir", cases.DirName, "verification cases directory")
	outDir := fs.String("out-dir", "", "directory for engine exports (never the case golden)")
	suite := fs.String("suite", "", "suite directory to run (default: all suites under --cases-dir)")
	var caseIDs []string
	fs.Func("case", "case id to run (suite/name; repeatable or comma-separated; default: cases with a golden)", func(s string) error {
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
		return fmt.Errorf("usage: calipers verify --engine <path|excel> [--recalculate] [--package] [--cases-dir DIR] [--out-dir DIR] [--suite NAME] [--case ID]...")
	}
	if fs.NArg() == 1 {
		*casesDir = fs.Arg(0)
	}
	host, err := hostFromSpec(*engineSpec, *recalculate)
	if err != nil {
		return err
	}
	policy := "host default (no recalculation requested)"
	if *recalculate {
		policy = "recalculate before export (external engine --recalculate)"
	}
	fmt.Fprintf(os.Stdout, "calculation policy: %s\n", policy)
	return runVerifyFilter(host, *casesDir, *suite, caseIDs, *outDir, os.Stdout, *packageDiag)
}

func hostFromSpec(spec string, recalculate bool) (engine, error) {
	spec = strings.TrimSpace(spec)
	switch spec {
	case "":
		return nil, fmt.Errorf("verify requires --engine <path|excel>")
	case "excel":
		if recalculate {
			return nil, fmt.Errorf("--recalculate is unsupported with --engine excel: the Excel host does not explicitly control calculation")
		}
		return newExcelHost(), nil
	default:
		return newBinaryHost(spec, recalculate), nil
	}
}

type caseOutcome struct {
	ID     string
	Export string
	Status string // pass, fail, error
	Detail string
}

func runVerify(eng engine, casesDir string, caseIDs []string, outDir string, w io.Writer) error {
	return runVerifyFilter(eng, casesDir, "", caseIDs, outDir, w, false)
}

func runVerifyFilter(eng engine, casesDir, suite string, caseIDs []string, outDir string, w io.Writer, packageDiag bool) error {
	corpus, err := cases.LoadCorpus(casesDir)
	if err != nil {
		return err
	}
	all, err := cases.FilterSuite(corpus.Cases, suite, corpus.Suites)
	if err != nil {
		return err
	}
	var selected []cases.Case
	if suite == "" && len(caseIDs) == 0 {
		selected = cases.GoldenComparePass(all)
	} else {
		selected, err = cases.Select(all, caseIDs)
		if err != nil {
			return err
		}
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

	var nPass, nFail, nErr, nSkip int
	outcomes := make([]caseOutcome, 0, len(selected))
	for i, c := range selected {
		o := verifyOne(eng, c, outDir, packageDiag)
		outcomes = append(outcomes, o)
		switch o.Status {
		case "pass":
			nPass++
		case "fail":
			nFail++
		case "skip":
			nSkip++
		default:
			nErr++
		}
		fmt.Fprintf(w, "[%d/%d] %s %s", i+1, len(selected), c.ID, strings.ToUpper(o.Status))
		if o.Detail != "" {
			fmt.Fprintf(w, " %s", o.Detail)
		}
		fmt.Fprintln(w)
	}
	if nSkip > 0 {
		fmt.Fprintf(w, "verify: %d pass, %d fail, %d error, %d skip\n", nPass, nFail, nErr, nSkip)
	} else {
		fmt.Fprintf(w, "verify: %d pass, %d fail, %d error\n", nPass, nFail, nErr)
	}
	if nFail+nErr > 0 {
		fmt.Fprint(w, formatDifferences(outcomes))
		return fmt.Errorf("%d failed, %d error", nFail, nErr)
	}
	return nil
}

func formatDifferences(outcomes []caseOutcome) string {
	var b strings.Builder
	b.WriteString("\nDifferences:\n")
	for _, o := range outcomes {
		if o.Status != "fail" && o.Status != "error" {
			continue
		}
		fmt.Fprintf(&b, "  %s %s\n", o.ID, strings.ToUpper(o.Status))
		for _, line := range strings.Split(o.Detail, "\n") {
			line = strings.TrimSpace(line)
			if line == "" || strings.HasPrefix(line, "(") {
				continue
			}
			fmt.Fprintf(&b, "    %s\n", line)
		}
	}
	return b.String()
}

func hasGolden(c cases.Case) bool {
	st, err := os.Stat(c.GoldenPath)
	return err == nil && st.Size() > 0
}

func verifyOne(eng engine, c cases.Case, outDir string, packageDiag bool) caseOutcome {
	exportPath := filepath.Join(outDir, filepath.FromSlash(c.ID)+".xlsx")
	if !hasGolden(c) {
		return caseOutcome{
			ID:     c.ID,
			Export: exportPath,
			Status: "skip",
			Detail: "(no golden.xlsx; generate with calipers excel-save or excel-run)",
		}
	}
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

	sem, err := xlsxmodel.CompareFiles(exportPath, c.GoldenPath)
	if err != nil {
		return caseOutcome{ID: c.ID, Export: exportPath, Status: "error", Detail: err.Error()}
	}
	detail := formatSemantic(sem)
	if packageDiag {
		detail += formatPackage(exportPath, c.GoldenPath)
	}
	if sem.Equal {
		return caseOutcome{ID: c.ID, Export: exportPath, Status: "pass", Detail: detail}
	}
	return caseOutcome{ID: c.ID, Export: exportPath, Status: "fail", Detail: detail}
}

func formatSemantic(sem xlsxmodel.Result) string {
	if sem.Equal {
		return "(semantic match)"
	}
	noun := "difference"
	if len(sem.Diffs) != 1 {
		noun = "differences"
	}
	var b strings.Builder
	fmt.Fprintf(&b, "(%d %s)", len(sem.Diffs), noun)
	for _, d := range sem.Diffs {
		b.WriteString("\n  ")
		if d.Location != "" {
			fmt.Fprintf(&b, "%s: %s %s", d.Axis, d.Location, d.Detail)
		} else {
			fmt.Fprintf(&b, "%s: %s", d.Axis, d.Detail)
		}
	}
	return b.String()
}

func formatPackage(exportPath, goldenPath string) string {
	pkg, err := compare.Files(exportPath, goldenPath)
	if err != nil {
		return "\n  package: " + err.Error()
	}
	if pkg.Equal {
		return "\n  package: match"
	}
	line := fmt.Sprintf("\n  package: %d differing parts", len(pkg.Diffs))
	if len(pkg.Diffs) > 0 {
		line += " " + pkg.Diffs[0].Part
	}
	return line
}
