//go:build !headless
// +build !headless

package gui

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"slices"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/bytesyntax/schedulehelper/internal/core"
)

type FileSelection struct {
	Label    *widget.Label
	Path     string
	FileName string
}

/*
================================================================================
Start GUI and setup all input elements
================================================================================
*/
func RunStandalone() {
	// Map to store file selections
	var fileSelections = map[string]*FileSelection{
		"Input File":    {},
		"Settings File": {},
		"Footer File":   {},
	}

	var generateBtn *widget.Button

	// Create main application winmdow
	myApp := app.NewWithID("github.com/bytesyntax/schedulehelper")
	mainWindow := myApp.NewWindow("Main App")
	mainWindow.Resize(fyne.NewSize(300, 200))

	// Ensure any child windows close when main closes
	childWindows := []*fyne.Window{}

	openFile := func(title string) {
		// Create a new window just for the file dialog
		childWin := myApp.NewWindow("Select " + title)
		childWin.Resize(fyne.NewSize(800, 600))

		// Add to child window list for cleanup
		childWindows = append(childWindows, &childWin)

		// File open dialog
		dialog.ShowFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil {
				log.Println("Error:", err)
				childWin.Close()
				return
			}
			if reader == nil {
				childWin.Close()
				return
			}

			filePath := reader.URI().Path()
			fileName := filepath.Base(filePath)

			fs := fileSelections[title]
			fs.Path = filePath
			fs.FileName = fileName
			fs.Label.SetText("Selected: " + fileName)

			log.Printf("%s selected: %s", title, filePath)

			childWin.Close()
		}, childWin)
		generateBtn.Enable()
		childWin.Show()
	}

	layout := container.NewVBox()
	// Label to describe the application
	titleLbl := widget.NewLabel("Select input files and click 'Generate Schedule' to proceed")
	titleLbl.Importance = widget.HighImportance
	layout.Add(titleLbl)

	// Open file buttons/labels
	for title, fs := range fileSelections {
		fs.Label = widget.NewLabel("No file selected")
		fs.Label.TextStyle = fyne.TextStyle{
			Monospace: true,
		}

		openBtn := widget.NewButton("Select "+title, func(t string) func() {
			return func() {
				openFile(t)
			}
		}(title))
		openBtn.Importance = widget.MediumImportance
		fs.Label.Importance = widget.LowImportance
		layout.Add(openBtn)
		layout.Add(fs.Label)
	}

	// 'Generate' button
	generateBtn = widget.NewButton("Generate Schedule", func() {
		log.Println("Generating schedule with selected files...")

		f1, err := os.Open(fileSelections["Input File"].Path)
		if err != nil {
			log.Println("Error opening input file:", err)
			generateBtn.Disable()
			return
		}
		defer f1.Close()

		f2, err := os.Open(fileSelections["Settings File"].Path)
		if err != nil {
			log.Println("Error opening settings file, skipping!")
		}
		defer f2.Close()

		f3, err := os.Open(fileSelections["Footer File"].Path)
		if err != nil {
			log.Println("Error opening footer file, skipping!")
		}
		defer f3.Close()

		fileData, err := core.ProcessFiles(f1, f2, f3) // Call the function to generate the schedules
		if err != nil {
			log.Println("Error processing files:", err)
			dialog.ShowError(err, mainWindow)
			return
		}

		outputFolder := filepath.Dir(fileSelections["Input File"].Path)
		fileSummary := []string{}
		for fileName, fileContent := range fileData {
			filePath := fmt.Sprintf("%s/%s", outputFolder, fileName)
			fileSummary = append(fileSummary, fileName)
			err = os.WriteFile(filePath, fileContent, 0644)
			if err != nil {
				log.Println("Error writing file:", filePath, err)
				dialog.ShowError(err, mainWindow)
				return
			}
			log.Println("File written successfully:", filePath)
		}
		slices.Sort(fileSummary)
		dialog.ShowInformation("Files Created", fmt.Sprintf("File created successfully in: %s\n%s", outputFolder, strings.Join(fileSummary, ", ")), mainWindow)
	})
	generateBtn.Importance = widget.HighImportance
	generateBtn.Disable()
	layout.Add(generateBtn)

	mainWindow.SetContent(layout)

	// Ensure child windows close when the main window is closed
	mainWindow.SetOnClosed(func() {
		for _, w := range childWindows {
			if *w != nil {
				(*w).Close()
			}
		}
	})

	mainWindow.ShowAndRun()
}
