package excel

import (
	"encoding/json"
	"fmt"
	"io"
	"io/fs"
	"net/http"
	"strings"
	"sync"
	"time"
)

// Job is the payload the sideloaded add-in fetches from GET /job.
// The add-in Excel.runs Job.Script (corpus Office.js, including top-level await).
type Job struct {
	Script string `json:"script"`
}

// Done is POST /done from the add-in after Excel.run finishes or fails.
type Done struct {
	OK    bool   `json:"ok"`
	Error string `json:"error,omitempty"`
}

// EncodeJob serializes script text for the add-in.
func EncodeJob(script string) ([]byte, error) {
	return json.Marshal(Job{Script: script})
}

// ParseDone decodes the add-in's completion payload.
func ParseDone(body []byte) (Done, error) {
	var d Done
	if err := json.Unmarshal(body, &d); err != nil {
		return Done{}, fmt.Errorf("office.js done: %w", err)
	}
	return d, nil
}

// jobServer serves the sideload add-in (Office.js page) and the job handshake.
type jobServer struct {
	script string
	mu     sync.Mutex
	done   chan Done
	once   sync.Once
}

func newJobServer(script string) *jobServer {
	return &jobServer{
		script: script,
		done:   make(chan Done, 1),
	}
}

func (s *jobServer) complete(d Done) {
	s.once.Do(func() { s.done <- d })
}

func (s *jobServer) wait(timeout time.Duration) (Done, error) {
	t := time.NewTimer(timeout)
	defer t.Stop()
	select {
	case d := <-s.done:
		if !d.OK {
			msg := d.Error
			if msg == "" {
				msg = "office.js failed"
			}
			return d, fmt.Errorf("office.js: %s", msg)
		}
		return d, nil
	case <-t.C:
		return Done{}, fmt.Errorf("excel-run timed out waiting for Office.js after %s", timeout)
	}
}

func (s *jobServer) ServeHTTP(w http.ResponseWriter, r *http.Request) {
	switch {
	case r.Method == http.MethodGet && r.URL.Path == "/job":
		data, err := EncodeJob(s.script)
		if err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write(data)
	case r.Method == http.MethodPost && r.URL.Path == "/done":
		body, err := io.ReadAll(r.Body)
		if err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		d, err := ParseDone(body)
		if err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}
		s.complete(d)
		w.WriteHeader(http.StatusNoContent)
	default:
		serveAddinFile(w, r)
	}
}

func serveAddinFile(w http.ResponseWriter, r *http.Request) {
	path := strings.TrimPrefix(r.URL.Path, "/")
	if path == "" {
		path = "taskpane.html"
	}
	data, err := fs.ReadFile(webFS, "web/"+path)
	if err != nil {
		http.NotFound(w, r)
		return
	}
	switch {
	case strings.HasSuffix(path, ".html"):
		w.Header().Set("Content-Type", "text/html; charset=utf-8")
	case strings.HasSuffix(path, ".js"):
		w.Header().Set("Content-Type", "application/javascript; charset=utf-8")
	case strings.HasSuffix(path, ".png"):
		w.Header().Set("Content-Type", "image/png")
	case strings.HasSuffix(path, ".xml") || strings.HasSuffix(path, ".tmpl"):
		w.Header().Set("Content-Type", "application/xml")
	}
	_, _ = w.Write(data)
}
