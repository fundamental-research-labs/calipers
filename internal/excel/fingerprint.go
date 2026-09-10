package excel

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

const (
	// WindowsExcelApplication is the Application string written by desktop
	// Windows Excel. Macintosh Excel, Excel Online, and third-party writers
	// use other values.
	WindowsExcelApplication = "Microsoft Excel"
	// windowsExcel16Prefix is the AppVersion prefix for Excel 2016+/365.
	windowsExcel16Prefix = "16."
)

// CheckWindowsExcel16Export reports whether path is a ZIP workbook whose
// docProps/app.xml Application is exactly "Microsoft Excel" and AppVersion
// starts with "16." (Windows Excel 2016+/365). Macintosh Excel, Excel Online,
// Openpyxl, LibreOffice, node-xlsx-stream, ODF Converter, missing Application,
// and older AppVersion (12/14/15) fail.
func CheckWindowsExcel16Export(path string) error {
	app, ver, err := ReadAppProperties(path)
	if err != nil {
		return err
	}
	if app != WindowsExcelApplication {
		if app == "" {
			return fmt.Errorf("%s: missing Application in docProps/app.xml", path)
		}
		return fmt.Errorf("%s: Application %q, want %q", path, app, WindowsExcelApplication)
	}
	if !strings.HasPrefix(ver, windowsExcel16Prefix) {
		if ver == "" {
			return fmt.Errorf("%s: missing AppVersion in docProps/app.xml", path)
		}
		return fmt.Errorf("%s: AppVersion %q, want 16.x", path, ver)
	}
	return nil
}

// ReadAppProperties returns Application and AppVersion from path's
// docProps/app.xml. The file must be a ZIP workbook that contains that part.
func ReadAppProperties(path string) (application, appVersion string, err error) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		return "", "", fmt.Errorf("%s: not a zip workbook: %w", path, err)
	}
	defer zr.Close()

	var data []byte
	for _, f := range zr.File {
		if f.Name != "docProps/app.xml" {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			return "", "", fmt.Errorf("%s: open docProps/app.xml: %w", path, err)
		}
		data, err = io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return "", "", fmt.Errorf("%s: read docProps/app.xml: %w", path, err)
		}
		break
	}
	if data == nil {
		return "", "", fmt.Errorf("%s: missing docProps/app.xml", path)
	}
	application, appVersion, err = parseAppXML(data)
	if err != nil {
		return "", "", fmt.Errorf("%s: parse docProps/app.xml: %w", path, err)
	}
	return application, appVersion, nil
}

func parseAppXML(data []byte) (application, appVersion string, err error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", "", err
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "Application":
			var s string
			if err := dec.DecodeElement(&s, &se); err != nil {
				return "", "", err
			}
			application = strings.TrimSpace(s)
		case "AppVersion":
			var s string
			if err := dec.DecodeElement(&s, &se); err != nil {
				return "", "", err
			}
			appVersion = strings.TrimSpace(s)
		}
	}
	return application, appVersion, nil
}
