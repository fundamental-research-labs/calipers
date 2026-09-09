package main

import (
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"strings"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/compare"
	enginehost "github.com/fundamental-research-labs/calipers/internal/engine"
	"github.com/fundamental-research-labs/calipers/internal/excel"
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

const verifyUsage = `calipers verify --engine <path|excel> [--recalculate] [--cases-dir DIR] [--out-dir DIR] [--suite NAME] [--case ID]...

  For each case: run the engine (load init.xlsx, Office.js if present and
  non-empty), export a result xlsx (not the golden), compare its ZIP parts
  to the committed Excel golden. Counts are differing parts, not defects.

  --engine excel    Excel COM host (Windows)
  --engine PATH     external binary:  save <in> <out>
                                      run  <in> <script.js> <out>
  --recalculate     request full recalculation before export by passing
                    --recalculate after save/run; requires engine support.
                    Unsupported with --engine excel. Default: host policy
                    (Mog preserves imported caches; Excel controls calculation).
  --suite NAME      run only this suite directory under --cases-dir
                    (e.g. roundtrip, default, scratch)
  --case ID         run only this case (repeatable or comma-separated;
                    id is suite/name, e.g. roundtrip/tier_a_simple)

  Default walk is committed-golden suites (roundtrip and default),
  tier_a and tier_b (skip tier_c hostiles). scratch is Office.js from
  an empty init with no committed goldens; use --suite scratch.
  --engine may be omitted when MOG_BIN or vendor/mog CLI artefact is set.
`

func verifyCmd(args []string) error {
	fs := flag.NewFlagSet("verify", flag.ContinueOnError)
	fs.SetOutput(os.Stderr)
	engineSpec := fs.String("engine", "", "engine binary path, or 'excel'")
	recalculate := fs.Bool("recalculate", false, "request recalculation before export (external engines supporting --recalculate only)")
	casesDir := fs.String("cases-dir", cases.DirName, "verification cases directory")
	outDir := fs.String("out-dir", "", "directory for engine exports (never the case golden)")
	suite := fs.String("suite", "", "suite directory to run (default: all suites under --cases-dir)")
	var caseIDs []string
	fs.Func("case", "case id to run (suite/name; repeatable or comma-separated; default: tier_a and tier_b)", func(s string) error {
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
		return fmt.Errorf("usage: calipers verify --engine <path|excel> [--recalculate] [--cases-dir DIR] [--out-dir DIR] [--suite NAME] [--case ID]...")
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
	return runVerifyFilter(host, *casesDir, *suite, caseIDs, *outDir, os.Stdout)
}

func hostFromSpec(spec string, recalculate bool) (engine, error) {
	spec = strings.TrimSpace(spec)
	if spec == "" {
		spec = defaultEngineSpec()
	}
	switch spec {
	case "":
		return nil, fmt.Errorf("verify requires --engine <path|excel> (or MOG_BIN / vendor/mog artefact)")
	case "excel":
		if recalculate {
			return nil, fmt.Errorf("--recalculate is unsupported with --engine excel: the Excel host does not explicitly control calculation")
		}
		return newExcelHost(), nil
	default:
		return newBinaryHost(spec, recalculate), nil
	}
}

func defaultEngineSpec() string {
	if b := os.Getenv("MOG_BIN"); b != "" {
		return b
	}
	if b := os.Getenv("ENGINE"); b != "" {
		return b
	}
	return findVendorMog()
}

func findVendorMog() string {
	name := "mog"
	if runtime.GOOS == "windows" {
		name = "mog.exe"
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
			for _, p := range []string{
				filepath.Join(dir, "vendor", "mog", "target-native", "debug", name),
				filepath.Join(dir, "vendor", "mog", "target-native", "release", name),
			} {
				if st, err := os.Stat(p); err == nil && !st.IsDir() {
					return p
				}
			}
			parent := filepath.Dir(dir)
			if parent == dir {
				break
			}
		}
	}
	return ""
}

type caseOutcome struct {
	ID     string
	Export string
	Status string // pass, fail, error
	Detail string
}

func runVerify(eng engine, casesDir string, caseIDs []string, outDir string, w io.Writer) error {
	return runVerifyFilter(eng, casesDir, "", caseIDs, outDir, w)
}

func runVerifyFilter(eng engine, casesDir, suite string, caseIDs []string, outDir string, w io.Writer) error {
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
	exportPath := filepath.Join(outDir, filepath.FromSlash(c.ID)+".xlsx")
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
		return caseOutcome{ID: c.ID, Export: exportPath, Status: "pass", Detail: "(package match)"}
	}
	detail := fmt.Sprintf("(differing package parts: %d)", len(got.Diffs))
	if len(got.Diffs) > 0 {
		detail += " " + got.Diffs[0].Part
	}
	return caseOutcome{ID: c.ID, Export: exportPath, Status: "fail", Detail: detail}
}
