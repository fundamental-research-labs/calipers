package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"github.com/fundamental-research-labs/calipers/internal/golden"
	"io"
	"os"
	"path/filepath"
	"testing"
)

func TestCommittedInputsHaveFormulasWithoutCaches(t *testing.T) {
	for _, s := range specs() {
		t.Run(s.name, func(t *testing.T) {
			got, err := workbook(s)
			if err != nil {
				t.Fatal(err)
			}
			committed, err := os.ReadFile(filepath.Join("../../verification/cases/recalculate", s.name, "init.xlsx"))
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(got, committed) {
				t.Fatal("regenerate inputs with go run ./scripts/gen-recalculate-cases")
			}
			zr, err := zip.NewReader(bytes.NewReader(got), int64(len(got)))
			if err != nil {
				t.Fatal(err)
			}
			count := 0
			for _, f := range zr.File {
				if f.Name != "xl/worksheets/sheet1.xml" {
					continue
				}
				r, err := f.Open()
				if err != nil {
					t.Fatal(err)
				}
				data, err := io.ReadAll(r)
				r.Close()
				if err != nil {
					t.Fatal(err)
				}
				var sheet struct {
					Rows []struct {
						Cells []struct {
							Formula *string `xml:"f"`
							Value   *string `xml:"v"`
							Inline  *string `xml:"is"`
						} `xml:"c"`
					} `xml:"sheetData>row"`
				}
				if err := xml.Unmarshal(data, &sheet); err != nil {
					t.Fatal(err)
				}
				for _, row := range sheet.Rows {
					for _, c := range row.Cells {
						if c.Formula != nil {
							count++
							if c.Value != nil || c.Inline != nil {
								t.Fatal("formula has a result cache")
							}
						}
					}
				}
			}
			if count != len(s.formulas) {
				t.Fatalf("%d formulas, want %d", count, len(s.formulas))
			}
			// A workbook that was merely copied must never pass the golden-value check.
			if err := checkGolden(filepath.Join("../../verification/cases/recalculate", s.name, "init.xlsx"), s); err == nil {
				t.Fatal("accepted uncalculated input")
			}
		})
	}
}

func TestWindowsGoldensWhenPresent(t *testing.T) {
	for _, s := range specs() {
		path := filepath.Join("../../verification/cases/recalculate", s.name, "golden.xlsx")
		if _, err := os.Stat(path); os.IsNotExist(err) {
			continue
		}
		if err := checkGolden(path, s); err != nil {
			t.Error(err)
		}
		meta, err := golden.Read(path)
		if err != nil || meta.Host != "excel-win" || meta.OS != "windows" || !meta.Recalculate || meta.Script != "" || meta.ExcelVersion == "" || meta.ExcelBuild == "" {
			t.Errorf("%s: invalid recalculation provenance: %+v, %v", s.name, meta, err)
		}
	}
}
