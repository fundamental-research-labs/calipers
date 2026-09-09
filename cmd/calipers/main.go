// calipers drives Excel (Windows COM) to produce golden xlsx files
// and verifies mog exports against those goldens.
//
//	calipers excel-save <input.xlsx> <output.xlsx>
//	calipers excel-run <input.xlsx> <script.js> <output.xlsx>
//	calipers verify --engine <path|excel> [--suite NAME] [--case ID]...
//	calipers version
package main

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
	"github.com/fundamental-research-labs/calipers/internal/golden"
)

const toolVersion = "0.1.0"

func main() {
	if err := run(os.Args[1:]); err != nil {
		fmt.Fprintf(os.Stderr, "calipers: %v\n", err)
		os.Exit(1)
	}
}

func run(args []string) error {
	if len(args) == 0 || args[0] == "-h" || args[0] == "--help" {
		printUsage()
		return nil
	}

	switch args[0] {
	case "excel-save":
		if len(args) != 3 {
			return fmt.Errorf("usage: calipers excel-save <input.xlsx> <output.xlsx>")
		}
		return excelSave(args[1], args[2])
	case "excel-save-pass":
		root := cases.DirName
		switch len(args) {
		case 1:
			// default verification/cases
		case 2:
			root = args[1]
		default:
			return fmt.Errorf("usage: calipers excel-save-pass [cases-dir]")
		}
		return excelSavePass(root)
	case "excel-run":
		if len(args) != 4 {
			return fmt.Errorf("usage: calipers excel-run <input.xlsx> <script.js> <output.xlsx>")
		}
		return excelRun(args[1], args[2], args[3])
	case "verify":
		return verifyCmd(args[1:])
	case "version":
		return printVersion()
	default:
		return fmt.Errorf("unknown command %q\n\n%s", args[0], usage)
	}
}

const usage = `calipers — generate Excel goldens (Windows COM)

Commands:
  excel-save <input.xlsx> <output.xlsx>
      Open input in Excel and Save As xlsx to output (load+save, no Office.js).
      Requires Windows + Microsoft Excel. Off Windows this command errors.

  excel-save-pass [cases-dir]
      Generate goldens for the default pass (tier_a and tier_b, skip
      Office.js and tier_c) via the same path as excel-save / excel-run.
      Default cases-dir is verification/cases.

  excel-run <input.xlsx> <script.js> <output.xlsx>
      Open input in Excel, run Office.js inside Excel (sideloaded add-in),
      then Save As xlsx. Requires Windows + Excel. Off Windows this errors.

  verify --engine <path|excel> [--recalculate] [--cases-dir DIR] [--out-dir DIR] [--suite NAME] [--case ID]...
      Run each case in an engine (Excel or an external binary) and
      compare export ZIP parts to the committed Excel golden.
      Binary argv: save <in.xlsx> <out.xlsx>
                   run  <in.xlsx> <script.js> <out.xlsx>
      Default walk is every suite under verification/cases (roundtrip
      and default), tier_a and tier_b (skip tier_c). --suite NAME runs
      one suite directory; --case suite/name runs one test.

  version
      Print calipers version. On Windows, also print Excel version
      when Excel.Application is registered.

  -h, --help
      Show this help.

Goldens must be generated on Windows. macOS Excel is not used.
`

func printUsage() {
	fmt.Fprint(os.Stdout, usage)
}

func excelSave(inputPath, outputPath string) error {
	return generateGolden(excel.NewHost(), inputPath, "", outputPath)
}

func excelRun(inputPath, scriptPath, outputPath string) error {
	return generateGolden(excel.NewHost(), inputPath, scriptPath, outputPath)
}

// generateGolden is the single Excel-win golden path: open the init, run
// Office.js when scriptPath is non-empty, Save As, write sidecar.
func generateGolden(host excel.Host, inputPath, scriptPath, outputPath string) error {
	var err error
	if scriptPath != "" {
		err = host.RunScript(inputPath, scriptPath, outputPath)
	} else {
		err = host.OpenSave(inputPath, outputPath)
	}
	if err != nil {
		return err
	}
	info, err := host.Info()
	if err != nil {
		return fmt.Errorf("excel info after save: %w", err)
	}
	if info.ID != excel.HostID {
		return fmt.Errorf("refuse golden host %q (want %s)", info.ID, excel.HostID)
	}
	meta := sidecarMeta(info, inputPath, scriptPath)
	if err := golden.Write(outputPath, meta); err != nil {
		return err
	}
	printf("wrote %s\n", outputPath)
	printf("wrote %s\n", golden.PathFor(outputPath))
	return nil
}

func printf(format string, args ...any) {
	fmt.Printf(format, args...)
	_ = os.Stdout.Sync()
}

func sidecarMeta(info excel.HostInfo, inputPath, scriptPath string) golden.Meta {
	m := golden.New(info.ID, info.OS, info.ExcelVersion, info.ExcelBuild, filepath.Base(inputPath))
	if scriptPath != "" {
		m.Script = filepath.Base(scriptPath)
	}
	return m
}

func printVersion() error {
	fmt.Printf("calipers %s\n", toolVersion)
	host := excel.NewHost()
	if !host.Available() {
		return nil
	}
	info, err := host.Info()
	if err != nil {
		fmt.Printf("excel: %v\n", err)
		return nil
	}
	fmt.Printf("excel host=%s version=%s build=%s os=%s\n",
		info.ID, info.ExcelVersion, info.ExcelBuild, info.OS)
	return nil
}
