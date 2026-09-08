package excel

import (
	"bytes"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
	"time"
)

func TestEncodeJobHandsScriptToAddin(t *testing.T) {
	script, err := os.ReadFile(corpusScript(t))
	if err != nil {
		t.Fatal(err)
	}
	raw, err := EncodeJob(string(script))
	if err != nil {
		t.Fatal(err)
	}
	var job Job
	if err := json.Unmarshal(raw, &job); err != nil {
		t.Fatal(err)
	}
	if job.Script != string(script) {
		t.Fatal("EncodeJob must pass corpus script through unchanged")
	}
	if !bytes.Contains(raw, []byte("Excel.run")) {
		t.Fatalf("job JSON missing Excel.run:\n%s", raw)
	}
}

func TestParseDone(t *testing.T) {
	ok, err := ParseDone([]byte(`{"ok":true}`))
	if err != nil || !ok.OK {
		t.Fatalf("ok: %+v %v", ok, err)
	}
	fail, err := ParseDone([]byte(`{"ok":false,"error":"boom"}`))
	if err != nil || fail.OK || fail.Error != "boom" {
		t.Fatalf("fail: %+v %v", fail, err)
	}
}

func TestJobServerHandshake(t *testing.T) {
	script := "await Excel.run(async (context) => { context.workbook.getActiveCell(); });"
	srv := newJobServer(script)
	ts := httptest.NewServer(srv)
	defer ts.Close()

	res, err := http.Get(ts.URL + "/job")
	if err != nil {
		t.Fatal(err)
	}
	body, err := io.ReadAll(res.Body)
	_ = res.Body.Close()
	if err != nil {
		t.Fatal(err)
	}
	var job Job
	if err := json.Unmarshal(body, &job); err != nil {
		t.Fatal(err)
	}
	if job.Script != script {
		t.Fatalf("GET /job script = %q", job.Script)
	}

	page, err := http.Get(ts.URL + "/taskpane.html")
	if err != nil {
		t.Fatal(err)
	}
	html, err := io.ReadAll(page.Body)
	_ = page.Body.Close()
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(html, []byte("office.js")) || !bytes.Contains(html, []byte("taskpane.js")) {
		t.Fatalf("taskpane.html must load Office.js, got:\n%s", html)
	}

	jsRes, err := http.Get(ts.URL + "/taskpane.js")
	if err != nil {
		t.Fatal(err)
	}
	js, err := io.ReadAll(jsRes.Body)
	_ = jsRes.Body.Close()
	if err != nil {
		t.Fatal(err)
	}
	for _, need := range []string{"Excel.run", "/job", "/done", "AsyncFunction"} {
		if !bytes.Contains(js, []byte(need)) {
			t.Fatalf("taskpane.js missing %q", need)
		}
	}

	waitErr := make(chan error, 1)
	go func() {
		_, err := srv.wait(2 * time.Second)
		waitErr <- err
	}()
	doneRes, err := http.Post(ts.URL+"/done", "application/json", strings.NewReader(`{"ok":true}`))
	if err != nil {
		t.Fatal(err)
	}
	_ = doneRes.Body.Close()
	if err := <-waitErr; err != nil {
		t.Fatalf("wait after ok: %v", err)
	}
}

func TestJobServerDoneError(t *testing.T) {
	srv := newJobServer("await Excel.run(async () => {});")
	ts := httptest.NewServer(srv)
	defer ts.Close()

	waitErr := make(chan error, 1)
	go func() {
		_, err := srv.wait(2 * time.Second)
		waitErr <- err
	}()
	_, err := http.Post(ts.URL+"/done", "application/json", strings.NewReader(`{"ok":false,"error":"range boom"}`))
	if err != nil {
		t.Fatal(err)
	}
	err = <-waitErr
	if err == nil || !strings.Contains(err.Error(), "range boom") {
		t.Fatalf("wait error = %v", err)
	}
}

func corpusScript(t *testing.T) string {
	t.Helper()
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("caller")
	}
	p := filepath.Join(filepath.Dir(file), "..", "..", "verification", "cases", "tier_a_simple_set_a1", "script.js")
	abs, err := filepath.Abs(p)
	if err != nil {
		t.Fatal(err)
	}
	return abs
}
