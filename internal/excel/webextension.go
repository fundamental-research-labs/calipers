package excel

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"strings"
)

const (
	contentTypesName = "[Content_Types].xml"
	workbookRelsName = "xl/_rels/workbook.xml.rels"
	taskpanesName    = "xl/webextensions/taskpanes.xml"
	taskpanesRels    = "xl/webextensions/_rels/taskpanes.xml.rels"
	webextName       = "xl/webextensions/webextension1.xml"
)

// StampWebExtension copies src xlsx to dst and adds a task-pane web extension
// that auto-opens the Calipers add-in (Office.AutoShowTaskpaneWithDocument).
func StampWebExtension(dst, src, addinID string) error {
	if addinID == "" {
		addinID = AddinID
	}
	r, err := zip.OpenReader(src)
	if err != nil {
		return fmt.Errorf("stamp webextension: %w", err)
	}
	defer r.Close()

	var buf bytes.Buffer
	w := zip.NewWriter(&buf)
	seen := map[string]bool{}
	for _, f := range r.File {
		seen[f.Name] = true
		rc, err := f.Open()
		if err != nil {
			_ = w.Close()
			return err
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			_ = w.Close()
			return err
		}
		switch f.Name {
		case contentTypesName:
			data = ensureContentTypes(data)
		case workbookRelsName:
			data = ensureWorkbookRels(data)
		}
		if err := writeZip(w, f.Name, data); err != nil {
			_ = w.Close()
			return err
		}
	}
	parts := map[string][]byte{
		taskpanesName: []byte(taskpanesXML),
		taskpanesRels: []byte(taskpanesRelsXML),
		webextName:    []byte(webextensionXML(addinID)),
	}
	if !seen[workbookRelsName] {
		if err := writeZip(w, workbookRelsName, ensureWorkbookRels([]byte(relsRoot))); err != nil {
			_ = w.Close()
			return err
		}
	}
	for name, data := range parts {
		if seen[name] {
			continue
		}
		if err := writeZip(w, name, data); err != nil {
			_ = w.Close()
			return err
		}
	}
	if err := w.Close(); err != nil {
		return err
	}
	return os.WriteFile(dst, buf.Bytes(), 0o644)
}

func writeZip(w *zip.Writer, name string, data []byte) error {
	fw, err := w.Create(name)
	if err != nil {
		return err
	}
	_, err = fw.Write(data)
	return err
}

func ensureContentTypes(data []byte) []byte {
	s := string(data)
	need := []string{
		`<Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>`,
		`<Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>`,
	}
	for _, n := range need {
		if !strings.Contains(s, n) {
			s = strings.Replace(s, "</Types>", n+"</Types>", 1)
		}
	}
	return []byte(s)
}

const relsRoot = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`

func ensureWorkbookRels(data []byte) []byte {
	s := string(data)
	const rel = `<Relationship Id="rIdCalipersWE" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="webextensions/taskpanes.xml"/>`
	if strings.Contains(s, "webextensiontaskpanes") {
		return data
	}
	if strings.Contains(s, "</Relationships>") {
		s = strings.Replace(s, "</Relationships>", rel+"</Relationships>", 1)
		return []byte(s)
	}
	return []byte(strings.TrimSuffix(relsRoot, "</Relationships>") + rel + "</Relationships>")
}

const taskpanesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:taskpane dockstate="right" visibility="1" width="350" row="0">
    <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
  </wetp:taskpane>
</wetp:taskpanes>
`

const taskpanesRelsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>
`

func webextensionXML(id string) string {
	return fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{%s}">
  <we:reference id="%s" version="1.0.0.0" store="developer" storeType="Registry"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
`, id, id)
}
