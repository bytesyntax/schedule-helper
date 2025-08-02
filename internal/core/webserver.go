package core

import (
	"archive/zip"
	"bytes"
	"fmt"
	"net/http"
)

/*
================================================================================
Start web server
================================================================================
*/
func RunHeadless() {
	http.HandleFunc("/", uploadHandler)
	fmt.Println("Server started at http://localhost:8999")
	http.ListenAndServe(":8999", nil)
}

/*
================================================================================
Handle web requests in containerized environment
================================================================================
*/
func uploadHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != "POST" {
		http.ServeFile(w, r, "upload.html")
		return
	}

	// Parse multipart form
	err := r.ParseMultipartForm(10 << 20) // 10 MB
	if err != nil {
		http.Error(w, "Cannot parse form", http.StatusBadRequest)
		return
	}

	// Retrieve files
	inputFile, _, err := r.FormFile("inputFile")
	if err != nil {
		http.Error(w, "Required input file missing", http.StatusBadRequest)
		return
	}
	defer inputFile.Close()

	settingsFile, _, _ := r.FormFile("settingsFile")
	if settingsFile != nil {
		defer settingsFile.Close()
	}

	footerFile, _, _ := r.FormFile("footerFile")
	if footerFile != nil {
		defer footerFile.Close()
	}

	// Save or process the files
	result, err := ProcessFiles(inputFile, settingsFile, footerFile)
	if err != nil {
		http.Error(w, "Error processing files: "+err.Error(), http.StatusInternalServerError)
		return
	}

	zipAndReturnFiles(w, result)
	fmt.Fprintf(w, "Files received and processed")
}

/*
================================================================================
Zip and return processed files
================================================================================
*/
func zipAndReturnFiles(w http.ResponseWriter, files map[string][]byte) {
	var buf bytes.Buffer
	zipWriter := zip.NewWriter(&buf)

	for name, content := range files {
		f, _ := zipWriter.Create(name)
		f.Write(content)
	}
	zipWriter.Close()

	w.Header().Set("Content-Disposition", "attachment; filename=schedules.zip")
	w.Header().Set("Content-Type", "application/zip")
	w.Write(buf.Bytes())
}
