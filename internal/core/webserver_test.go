package core

import (
	"archive/zip"
	"bytes"
	"io"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestZipAndReturnFiles(t *testing.T) {
	files := map[string][]byte{
		"a.txt":      []byte("hello world"),
		"b/data.txt": []byte("other content"),
	}

	rr := httptest.NewRecorder()
	zipAndReturnFiles(rr, files)

	res := rr.Result()
	defer res.Body.Close()

	if res.Header.Get("Content-Type") != "application/zip" {
		t.Fatalf("expected application/zip content-type, got %s", res.Header.Get("Content-Type"))
	}
	if res.Header.Get("X-Processed") != "true" {
		t.Fatalf("expected X-Processed header true, got %s", res.Header.Get("X-Processed"))
	}

	body, err := io.ReadAll(res.Body)
	if err != nil {
		t.Fatalf("reading response body: %v", err)
	}
	if len(body) == 0 {
		t.Fatalf("expected non-empty zip body")
	}

	zr, err := zip.NewReader(bytes.NewReader(body), int64(len(body)))
	if err != nil {
		t.Fatalf("opening zip reader: %v", err)
	}

	found := map[string]bool{}
	for _, f := range zr.File {
		found[f.Name] = true
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("opening zip entry %s: %v", f.Name, err)
		}
		_, _ = io.ReadAll(rc)
		rc.Close()
	}

	for name := range files {
		if !found[name] {
			t.Fatalf("expected zip to contain %s, but it did not", name)
		}
	}
}

func TestUploadHandler_GETServesUploadHTML(t *testing.T) {
	// Create a temporary upload.html in repo root so ServeFile finds it.
	// Tests run with repo root as working directory.
	tmpName := "upload.html"
	content := "<html><body>upload page</body></html>"
	if err := os.WriteFile(tmpName, []byte(content), 0o644); err != nil {
		t.Fatalf("failed creating temporary upload.html: %v", err)
	}
	t.Cleanup(func() { os.Remove(tmpName) })

	req := httptest.NewRequest(http.MethodGet, "/", nil)
	rr := httptest.NewRecorder()
	uploadHandler(rr, req)

	res := rr.Result()
	defer res.Body.Close()

	if res.StatusCode != http.StatusOK {
		t.Fatalf("expected status 200, got %d", res.StatusCode)
	}

	b, err := io.ReadAll(res.Body)
	if err != nil {
		t.Fatalf("reading body: %v", err)
	}
	got := string(b)
	// Ensure basic content from our temp file is present
	if !strings.Contains(got, "upload page") {
		// if ServeFile resolved to a different path, include debugging hint
		t.Fatalf("response body does not contain expected content; length=%d; path=%s", len(got), filepath.Join(".", tmpName))
	}
}

func TestUploadHandler_POSTProcessesFiles(t *testing.T) {
	// Stub ProcessFunc
	old := ProcessFunc
	defer func() { ProcessFunc = old }()
	ProcessFunc = func(input, settings, footer io.Reader) (map[string][]byte, error) {
		// ensure input is passed
		b, err := io.ReadAll(input)
		if err != nil {
			t.Fatalf("reading input in stub: %v", err)
		}
		if len(b) == 0 {
			t.Fatalf("expected input file to have content")
		}
		return map[string][]byte{"generated.txt": []byte("ok")}, nil
	}

	var body bytes.Buffer
	mw := multipart.NewWriter(&body)
	// inputFile (required)
	fw, err := mw.CreateFormFile("inputFile", "in.xlsx")
	if err != nil {
		t.Fatalf("create form file: %v", err)
	}
	fw.Write([]byte("dummyinput"))
	// optional settings
	fw, _ = mw.CreateFormFile("settingsFile", "settings.xlsx")
	fw.Write([]byte("settings"))
	// optional footer
	fw, _ = mw.CreateFormFile("footerFile", "footer.xlsx")
	fw.Write([]byte("footer"))
	mw.Close()

	req := httptest.NewRequest(http.MethodPost, "/", &body)
	req.Header.Set("Content-Type", mw.FormDataContentType())
	rr := httptest.NewRecorder()
	uploadHandler(rr, req)

	res := rr.Result()
	defer res.Body.Close()

	if res.StatusCode != http.StatusOK {
		t.Fatalf("expected status 200, got %d", res.StatusCode)
	}
	if res.Header.Get("Content-Type") != "application/zip" {
		t.Fatalf("expected application/zip content-type, got %s", res.Header.Get("Content-Type"))
	}
	if res.Header.Get("X-Processed") != "true" {
		t.Fatalf("expected X-Processed header true, got %s", res.Header.Get("X-Processed"))
	}
}
