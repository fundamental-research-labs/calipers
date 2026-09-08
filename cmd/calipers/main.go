// calipers drives Excel (Windows COM) to produce golden xlsx files.
//
//	calipers excel-save <input.xlsx> <output.xlsx>
//	calipers excel-run <input.xlsx> <script.js> <output.xlsx>
//	calipers version
package main

import (
	"fmt"
	"os"
	"path/filepath"

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
	case "excel-run":
		if len(args) != 4 {
			return fmt.Errorf("usage: calipers excel-run <input.xlsx> <script.js> <output.xlsx>")
		}
		return excelRun(args[1], args[2], args[3])
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

  excel-run <input.xlsx> <script.js> <output.xlsx>
      Open input in Excel, run Office.js inside Excel (sideloaded add-in),
      then Save As xlsx. Requires Windows + Excel. Off Windows this errors.

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
	host := excel.NewHost()
	if err := host.OpenSave(inputPath, outputPath); err != nil {
		return err
	}
	info, err := host.Info()
	if err != nil {
		return fmt.Errorf("excel info after save: %w", err)
	}
	meta := sidecarMeta(info, inputPath, "")
	if err := golden.Write(outputPath, meta); err != nil {
		return err
	}
	fmt.Printf("wrote %s\n", outputPath)
	fmt.Printf("wrote %s\n", golden.PathFor(outputPath))
	return nil
}

func excelRun(inputPath, scriptPath, outputPath string) error {
	host := excel.NewHost()
	if err := host.RunScript(inputPath, scriptPath, outputPath); err != nil {
		return err
	}
	info, err := host.Info()
	if err != nil {
		return fmt.Errorf("excel info after run: %w", err)
	}
	meta := sidecarMeta(info, inputPath, scriptPath)
	if err := golden.Write(outputPath, meta); err != nil {
		return err
	}
	fmt.Printf("wrote %s\n", outputPath)
	fmt.Printf("wrote %s\n", golden.PathFor(outputPath))
	return nil
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
