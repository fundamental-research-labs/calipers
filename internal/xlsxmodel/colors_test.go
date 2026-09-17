package xlsxmodel

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestCanonicalColor(t *testing.T) {
	for _, tc := range []struct{ color, tint, want string }{
		{"FF4472C4", "", "4472C4"},
		{"4472c4", "0", "4472C4"},
		{"804472C4", "", "804472C4"},
		{"4472C4", "0.79998168889431442", "D9E1F2"},
		{"4472C4", "0.39997558519241921", "8EA9DB"},
		{"000000", "1", "FFFFFF"},
		{"FFFFFF", "-1", "000000"},
		{"theme:9", "0.4", "THEME:9@0.4000"},
	} {
		if got := canonicalColor(tc.color, tc.tint); got != tc.want {
			t.Errorf("%s tint %s: got %s, want %s", tc.color, tc.tint, got, tc.want)
		}
	}
}

func TestPatternedFillKeepsBackground(t *testing.T) {
	read := func(pattern, background string) string {
		raw := `<fill><patternFill patternType="` + pattern + `"><fgColor rgb="FFFF0000"/><bgColor rgb="` + background + `"/></patternFill></fill>`
		dec := xml.NewDecoder(strings.NewReader(raw))
		_, _ = dec.Token()
		return readFill(dec, nil)
	}
	if read("solid", "FF000000") != read("solid", "FFFFFFFF") {
		t.Fatal("solid background affected comparison")
	}
	if read("darkGrid", "FF000000") == read("darkGrid", "FFFFFFFF") {
		t.Fatal("patterned background difference was lost")
	}
	if canonicalColor("FF000000", "") == canonicalColor("FFFFFFFF", "") {
		t.Fatal("different foregrounds compared equal")
	}
}
