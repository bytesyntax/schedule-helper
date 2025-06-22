package main

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"io"
	"net/http"
	"slices"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/go-gota/gota/dataframe"
	"github.com/go-gota/gota/series"
	"github.com/xuri/excelize/v2"
)

type HourActivity int

const (
	StateFree HourActivity = iota
	StateWork
	StateLunch
	StateAssigned
)

type ShiftActivity struct {
	hourSchedule []HourActivity
	shiftTime    string
	employeeName string
	employeeId   int
	role         string
	phone        string
	shiftLength  string
}

type DaySchedule struct {
	headers []string
	dayStr  string
	dateStr string
	weekStr string
	shifts  []ShiftActivity
}

var (
	styleFree     = 0
	styleWork     = 0
	styleLunch    = 0
	styleAssigned = 0
	styleName     = 0
	styleTitle    = 0
	styleHeader   = 0
)

type FooterCell struct {
	Row   int
	Col   int
	Value interface{}
	Style *excelize.Style
	Merge string // optional: e.g., "C20:D20"
}

// func main() {
// Read settings file
// This file contains employeeId, phone and role
// settingsFile, err := excelize.OpenFile("Settings.xlsx")
// if err != nil {
// 	fmt.Println("Error opening settings file:", err)
// 	return
// }
// settingsData, err := settingsFile.GetRows(settingsFile.GetSheetName(settingsFile.GetActiveSheetIndex()))
// if err != nil {
// 	fmt.Println("Error getting settings rows:", err)
// 	return
// }
// if len(settingsData) < 2 {
// 	fmt.Println("Settings file is empty or has no data")
// 	return
// }
// settingsDf := excelRowsToDataFrame(settingsData[1:]) // Skip header row
// if settingsDf.Ncol() < 3 {
// 	fmt.Println("Settings file does not have enough columns")
// 	return
// }
// if settingsDf.Ncol() > 3 {
// 	fmt.Println("Settings file has more than 3 columns, only first 3 will be used")
// 	settingsDf = settingsDf.Subset([]int{0, 1, 2}) // Keep only first 3 columns
// }
// err = settingsDf.SetNames("employeeId", "phone", "role")
// if err != nil {
// 	fmt.Println("Error setting DataFrame column names:", err)
// 	return
// }

// Read input data
// This file contains employeeId, lastName, firstName, shiftType, date, time and department
// 	df, err := readAndRefineInputData("DailyStaffingSchedule_1348853_1494f5a70b167da50a8.xlsx", settingsDf)
// 	if err != nil {
// 		fmt.Println("Error reading input data:", err)
// 		return
// 	}

// 	// Read and prepare footer file
// 	// This file contains footer data and styles to be applied to each daily schedule
// 	footerFile, err := excelize.OpenFile("Footer.xlsx")
// 	if err != nil {
// 		fmt.Println("Error opening footer file:", err)
// 		return
// 	}
// 	footer, err := PrepareFooter(footerFile, footerFile.GetSheetName(footerFile.GetActiveSheetIndex()))
// 	if err != nil {
// 		fmt.Println("Error preparing footer:", err)
// 		return
// 	}

// 	// Create weekly schedules
// 	// This will create a new excel file for each week with daily schedules as sheets
// 	err = createWeekSchedules(df, footer)
// 	if err != nil {
// 		fmt.Println("Error creating weekly schedules", err)
// 		return
// 	}
// }

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
	result, err := processFiles(inputFile, settingsFile, footerFile)
	if err != nil {
		http.Error(w, "Error processing files: "+err.Error(), http.StatusInternalServerError)
		return
	}

	zipAndReturnFiles(w, result)
	fmt.Fprintf(w, "Files received and processed")
}

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

func processFiles(input io.Reader, settings io.Reader, footer io.Reader) (map[string][]byte, error) {
	// Here you would add your logic to read Excel files (e.g., using "github.com/xuri/excelize")
	// and generate outputs
	fmt.Println("Processing files...")
	var settingsDf dataframe.DataFrame
	settingsDf, _ = readSettingsFile(settings)
	df, err := readAndRefineInputData(input, settingsDf)
	if err != nil {
		return nil, errors.New("Error reading input data: " + err.Error())
	}
	footerData, _ := PrepareFooter(footer, "Sheet1")

	result, err := createWeekSchedules(df, footerData)
	if err != nil {
		return nil, errors.New("Error creating weekly schedules: " + err.Error())
	}

	return result, nil
}

func main() {
	http.HandleFunc("/", uploadHandler)
	fmt.Println("Server started at http://localhost:8999")
	http.ListenAndServe(":8999", nil)
}

func readSettingsFile(r io.Reader) (dataframe.DataFrame, error) {
	// Read settings file
	// This file contains employeeId, phone and role
	sr, err := excelize.OpenReader(r)
	if err != nil {
		return dataframe.DataFrame{}, errors.New("Error opening settings file: " + err.Error())
	}
	settingsData, err := sr.GetRows(sr.GetSheetName(sr.GetActiveSheetIndex()))
	if err != nil {
		return dataframe.DataFrame{}, errors.New("Error getting settings rows: " + err.Error())
	}
	if len(settingsData) < 2 {
		fmt.Println("Settings file is empty or has no data")
		return dataframe.DataFrame{}, nil
	}
	settingsDf := excelRowsToDataFrame(settingsData[1:]) // Skip header row
	if settingsDf.Ncol() < 3 {
		fmt.Println("Settings file does not have enough columns")
		return dataframe.DataFrame{}, nil
	}
	if settingsDf.Ncol() > 3 {
		fmt.Println("Settings file has more than 3 columns, only first 3 will be used")
		settingsDf = settingsDf.Subset([]int{0, 1, 2}) // Keep only first 3 columns
	}
	err = settingsDf.SetNames("employeeId", "phone", "role")
	if err != nil {
		fmt.Println("Error setting DataFrame column names:", err)
		return dataframe.DataFrame{}, nil
	}

	return settingsDf, nil
}

/*
================================================================================
Create a excel workbook per week
================================================================================
*/
func createWeekSchedules(df dataframe.DataFrame, footer []FooterCell) (map[string][]byte, error) {
	var results = make(map[string][]byte)

	var wg sync.WaitGroup
	wg.Add(len(df.GroupBy("weekNumber").GetGroups()))

	for weekNumber, weekDf := range df.GroupBy("weekNumber").GetGroups() {
		fmt.Printf("Processing weekNumber: %v\n", weekNumber)

		go func() {
			defer wg.Done()
			f := excelize.NewFile()
			setStyles(f)

			for _, dateDf := range weekDf.GroupBy("date").GetGroups() {
				err := createDaySchedule(f, dateDf, footer)
				if err != nil {
					panic(err)
				}
			}

			// Swap sort sheets, use auto created "Sheet1" as anchor (since move is only way to reorder sheets)
			sheetNames := f.GetSheetList()
			for _, day := range []string{"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"} {
				if slices.Contains(sheetNames, day) {
					f.MoveSheet(day, "Sheet1")
				}
			}
			f.DeleteSheet("Sheet1")

			fn := fmt.Sprintf("Vecka %v.xlsx", weekNumber)
			results[fn] = nil // Initialize the map entry
			// Save the file to a buffer
			var err error
			buf, err := f.WriteToBuffer()
			if err != nil {
				panic(fmt.Errorf("error writing to buffer: %v", err))
			}
			results[fn] = buf.Bytes()
			// f.SaveAs(fmt.Sprintf("Vecka %v.xlsx", weekNumber))
		}()
	}

	wg.Wait()
	return results, nil
}

/*
================================================================================
Create a sheet per day in the excel file
================================================================================
*/
func createDaySchedule(file *excelize.File, dateDf dataframe.DataFrame, footer []FooterCell) error {
	dateDf = dateDf.Arrange(
		dataframe.Sort("startTime"),
		dataframe.Sort("endTime"),
	)
	dayData, err := parseDayData(dateDf)
	if err != nil {
		return errors.New("error getting day schedule: " + err.Error())
	}

	sheetName := dayData.dayStr
	file.NewSheet(sheetName)

	totalOffset := 1
	// Title row
	titleStartCell, err := excelize.CoordinatesToCellName(1, 1)
	if err != nil {
		return fmt.Errorf("error getting title start cell: %v", err)
	}
	titleEndCell, err := excelize.CoordinatesToCellName(len(dayData.headers), 1)
	if err != nil {
		return fmt.Errorf("error getting title end cell: %v", err)
	}
	err = file.MergeCell(sheetName, titleStartCell, titleEndCell)
	if err != nil {
		return fmt.Errorf("error merging cells: %v", err)
	}
	file.SetCellValue(sheetName, titleStartCell, dayData.dayStr+" - "+dayData.dateStr)
	file.SetCellStyle(sheetName, titleStartCell, titleEndCell, styleTitle)
	totalOffset += 1

	// Header row
	for colIdx := 0; colIdx < len(dayData.headers); colIdx++ {
		cell, err := excelize.CoordinatesToCellName(colIdx+1, 2)
		if err != nil {
			return fmt.Errorf("error calculating cell for header row %v", err)
		}
		file.SetCellValue(sheetName, cell, dayData.headers[colIdx])
		file.SetCellStyle(sheetName, cell, cell, styleHeader)
	}
	totalOffset += 1

	// Shift rows (offset for previous rows
	rowOffset := totalOffset
	// First column with time data
	hourOffset := 4
	for dataIdx := 0; dataIdx < len(dayData.shifts); dataIdx += 1 {
		// Time col
		file.SetCellValue(sheetName, fmt.Sprintf("A%d", dataIdx+rowOffset), dayData.shifts[dataIdx].shiftTime)
		file.SetColWidth(sheetName, "A", "A", 15)
		// Name col
		file.SetCellValue(sheetName, fmt.Sprintf("B%d", dataIdx+rowOffset), dayData.shifts[dataIdx].employeeName)
		file.SetCellStyle(sheetName, fmt.Sprintf("B%d", dataIdx+rowOffset), fmt.Sprintf("B%d", dataIdx+rowOffset), styleName)
		file.SetColWidth(sheetName, "B", "B", 30)
		// Telephone col
		file.SetCellValue(sheetName, fmt.Sprintf("C%d", dataIdx+rowOffset), dayData.shifts[dataIdx].phone)
		// Time cols
		for hourIdx := 0; hourIdx < len(dayData.shifts[dataIdx].hourSchedule)-1; hourIdx++ { // Skip last hourSchedule since headers compacted by one!!!
			cell, err := excelize.CoordinatesToCellName(hourIdx+hourOffset, dataIdx+rowOffset)
			if err != nil {
				return fmt.Errorf("error calculating cell for hour %d, row %d: %v", hourIdx, dataIdx+rowOffset, err)
			}

			// For each hour for current row, set the value and style
			switch dayData.shifts[dataIdx].hourSchedule[hourIdx] {
			case StateFree:
				file.SetCellStyle(sheetName, cell, cell, styleFree)
			case StateWork:
				file.SetCellStyle(sheetName, cell, cell, styleWork)
			case StateLunch:
				file.SetCellValue(sheetName, cell, "Lunch")
				file.SetCellStyle(sheetName, cell, cell, styleLunch)
			case StateAssigned:
				file.SetCellValue(sheetName, cell, dayData.shifts[dataIdx].role)
				file.SetCellStyle(sheetName, cell, cell, styleAssigned)
			}
		}
		// Total time
		totalCol, err := excelize.CoordinatesToCellName(len(dayData.headers)+1, totalOffset)
		if err != nil {
			return errors.New("Error calculating totalCol" + err.Error())
		}
		file.SetCellValue(sheetName, totalCol, dayData.shifts[dataIdx].shiftLength)

		totalOffset += 1
	}
	timeColStart, err := excelize.ColumnNumberToName(hourOffset)
	if err != nil {
		return fmt.Errorf("error calculating time column start: %v", err)
	}
	timeColEnd, err := excelize.ColumnNumberToName(hourOffset + len(dayData.shifts[0].hourSchedule) - 1)
	if err != nil {
		return fmt.Errorf("error calculating time column end: %v", err)
	}
	file.SetColWidth(sheetName, timeColStart, timeColEnd, 12)

	// Header row (trailing)
	for colIdx := 0; colIdx < len(dayData.headers); colIdx++ {
		cell, err := excelize.CoordinatesToCellName(colIdx+1, totalOffset)
		if err != nil {
			return fmt.Errorf("error calculating cell for header row %v", err)
		}
		file.SetCellValue(sheetName, cell, dayData.headers[colIdx])
		file.SetCellStyle(sheetName, cell, cell, styleHeader)
	}
	totalOffset += 1

	ApplyFooterToSheet(file, sheetName, footer, totalOffset)

	return nil
}

/*
================================================================================
Read and refine input file with time data
================================================================================
*/
func readAndRefineInputData(r io.Reader, settingsDf dataframe.DataFrame) (dataframe.DataFrame, error) {
	fr, err := excelize.OpenReader(r)
	if err != nil {
		return dataframe.DataFrame{}, errors.New("Error opening file: " + err.Error())
	}

	rows, err := fr.GetRows("Worksheet")
	if err != nil {
		return dataframe.DataFrame{}, errors.New("Error getting rows: " + err.Error())
	}

	df := excelRowsToDataFrame(rows[1:])
	err = df.SetNames("employeeId", "lastName", "firstName", "shiftType", "date", "time", "department")
	if err != nil {
		return dataframe.DataFrame{}, errors.New("Error setting column names: " + err.Error())
	}

	df = df.Mutate(series.New(fixEmployeeId(df.Col("employeeId")), series.Int, "employeeId"))
	shiftSeries, err := extractShiftDetails(df.Col("time"))
	if err != nil {
		return dataframe.DataFrame{}, errors.New("Error extracting shift details: " + err.Error())
	}
	for _, s := range shiftSeries {
		df = df.Mutate(s)
	}

	settingsCols, err := extractSettingsCols(df.Col("employeeId"), settingsDf)
	if err != nil {
		return dataframe.DataFrame{}, errors.New("Error extracting shift roles: " + err.Error())
	}
	for _, s := range settingsCols {
		df = df.Mutate(s)
	}

	weekNumSeries, err := extractWeekNumber(df.Col("date"))
	if err != nil {
		return dataframe.DataFrame{}, errors.New("Error extracting week number: " + err.Error())
	}
	df = df.Mutate(weekNumSeries)
	df = df.Arrange(
		dataframe.Sort("date"),
	)

	return df, nil
}

func extractSettingsCols(s series.Series, dfSettings dataframe.DataFrame) ([]series.Series, error) {
	roleData := make([]string, len(s.Records()))
	phoneData := make([]string, len(s.Records()))
	for i, v := range s.Records() {
		employeeId, err := strconv.Atoi(v)
		if err != nil {
			return nil, errors.New("Error converting employeeId to int: " + err.Error())
		}

		employeeSettings := dfSettings.Filter(
			dataframe.F{
				Colname:    "employeeId",
				Comparator: series.Eq,
				Comparando: employeeId,
			},
		)
		if employeeSettings.Nrow() > 0 {
			employeeAssignment := employeeSettings.Col("role").Records()
			if len(employeeAssignment) > 0 {
				roleData[i] = employeeAssignment[0]
			}
			employeePhone := employeeSettings.Col("phone").Records()
			if len(employeePhone) > 0 {
				phoneData[i] = employeePhone[0]
			}
		}
	}

	return []series.Series{
		series.New(roleData, series.String, "role"),
		series.New(phoneData, series.String, "phone"),
	}, nil
}

// Convert Excel rows to a gota dataframe
func excelRowsToDataFrame(rows [][]string) dataframe.DataFrame {
	if len(rows) == 0 {
		return dataframe.DataFrame{}
	}
	columns := rows[0]
	records := rows[1:]

	// Ensure all rows have the same length as columns
	for i := range records {
		if len(records[i]) < len(columns) {
			// Pad with empty strings
			diff := len(columns) - len(records[i])
			for j := 0; j < diff; j++ {
				records[i] = append(records[i], "")
			}
		}
	}

	df := dataframe.LoadRecords(append([][]string{columns}, records...))
	return df
}

// Convert employeeId to int
func fixEmployeeId(s series.Series) series.Series {
	var convertVals = make([]int, len(s.Records()))
	for i, v := range s.Records() {
		val, err := strconv.Atoi(v)
		if err != nil {
			convertVals[i] = -1
		} else {
			convertVals[i] = val
		}
	}
	return series.New(convertVals, series.Int, s.Name) // Return a new series with the converted values
}

// Add extra columns based on shift time: start/end times, length and has lunch
func extractShiftDetails(s series.Series) ([]series.Series, error) {
	startTimes := make([]string, len(s.Records()))
	endTimes := make([]string, len(s.Records()))
	shiftLengths := make([]float64, len(s.Records()))
	hasLunch := make([]bool, len(s.Records()))

	for i, v := range s.Records() {
		tmp := strings.Split(v, " - ")
		if len(tmp) < 2 {
			return nil, errors.New("Invalid time format in record: " + v)
		}

		st, err := time.Parse("15:04", tmp[0])
		if err != nil {
			return nil, errors.New("Error parsing start time: " + err.Error())
		}
		startTimes[i] = st.Format(time.TimeOnly) // Format to a string representation

		en, err := time.Parse("15:04", tmp[1])
		if err != nil {
			return nil, errors.New("Error parsing end time: " + err.Error())
		}
		endTimes[i] = en.Format(time.TimeOnly)

		shiftLengths[i] = en.Sub(st).Hours()
		if shiftLengths[i] > 5 {
			hasLunch[i] = true
			shiftLengths[i] -= 1 // Subtract 1 hour for lunch if shift is longer than 5 hours
		}
	}
	return []series.Series{
		series.New(startTimes, series.String, "startTime"),
		series.New(endTimes, series.String, "endTime"),
		series.New(shiftLengths, series.Float, "shiftLength"),
		series.New(hasLunch, series.Bool, "hasLunch"),
	}, nil
}

// Add extra column based on date: week number
func extractWeekNumber(s series.Series) (series.Series, error) {
	weekNumbers := make([]int, len(s.Records()))
	for i, v := range s.Records() {
		date, err := time.Parse("2006-01-02", v)
		if err != nil {
			return series.New([]int{-1}, series.Int, "weekNumber"), errors.New("Invalid date format in record: " + v)
		}
		_, week := date.ISOWeek()
		weekNumbers[i] = week
	}
	return series.New(weekNumbers, series.Int, "weekNumber"), nil
}

/*
================================================================================
Produce data to be used in a day schedule sheet
================================================================================
*/
func parseDayData(df dataframe.DataFrame) (DaySchedule, error) {
	// Get earliest start
	dayStart, err := time.Parse(
		time.TimeOnly,
		df.Arrange(dataframe.Sort("startTime")).Col("startTime").Records()[0],
	)
	if err != nil {
		return DaySchedule{}, errors.New("Error parsing day start time: " + err.Error())
	}
	// Get latest end
	dayEnd, err := time.Parse(
		time.TimeOnly,
		df.Arrange(dataframe.RevSort("endTime")).Col("endTime").Records()[0],
	)
	if err != nil {
		return DaySchedule{}, errors.New("Error parsing day end time: " + err.Error())
	}
	// Generate hour slots
	hideBefore, err := time.Parse(time.TimeOnly, "10:00:00")
	if err != nil {
		return DaySchedule{}, errors.New("Error parsing hideBefore time: " + err.Error())
	}
	timeSlots := []time.Time{}
	for t := dayStart; !t.After(dayEnd); t = t.Add(time.Hour) {
		if t.Before(hideBefore) {
			continue
		}
		timeSlots = append(timeSlots, t)
	}
	if dayEnd.After(timeSlots[len(timeSlots)-1]) {
		timeSlots = append(timeSlots, dayEnd)
	}

	// Populate rows
	shiftRows := make([]ShiftActivity, df.Nrow())
	for rowIdx := 0; rowIdx < len(shiftRows); rowIdx++ {
		start, err := time.Parse(time.TimeOnly, df.Col("startTime").Elem(rowIdx).String())
		if err != nil {
			return DaySchedule{}, errors.New("Error parsing startTime: " + err.Error())
		}
		end, err := time.Parse(time.TimeOnly, df.Col("endTime").Elem(rowIdx).String())
		if err != nil {
			return DaySchedule{}, errors.New("Error parsing endTime: " + err.Error())
		}
		hasLunch, err := df.Col("hasLunch").Elem(rowIdx).Bool()
		if err != nil {
			return DaySchedule{}, errors.New("Error parsing hasLunch: " + err.Error())
		}
		employeeId, err := strconv.Atoi(df.Col("employeeId").Elem(rowIdx).String())
		if err != nil {
			return DaySchedule{}, errors.New("Error parsing employeeId: " + err.Error())
		}
		shiftRows[rowIdx] = ShiftActivity{
			shiftTime:    df.Col("time").Elem(rowIdx).String(),
			employeeName: df.Col("firstName").Elem(rowIdx).String() + " " + df.Col("lastName").Elem(rowIdx).String(),
			employeeId:   employeeId,
			shiftLength:  df.Col("shiftLength").Elem(rowIdx).String(),
			role:         df.Col("role").Elem(rowIdx).String(),
			phone:        df.Col("phone").Elem(rowIdx).String(),
			hourSchedule: make([]HourActivity, len(timeSlots)),
		}

		for hour, slot := range timeSlots {
			if slot.Before(start) {
				shiftRows[rowIdx].hourSchedule[hour] = StateFree
			} else if !slot.Before(start) && slot.Before(end) {
				// Lunch break after 5 hours
				if hasLunch && slot == start.Add(5*time.Hour) {
					shiftRows[rowIdx].hourSchedule[hour] = StateLunch
				} else if shiftRows[rowIdx].role != "" {
					shiftRows[rowIdx].hourSchedule[hour] = StateAssigned
				} else {
					shiftRows[rowIdx].hourSchedule[hour] = StateWork
				}
			} else {
				shiftRows[rowIdx].hourSchedule[hour] = StateFree
			}
		}
	}

	// Compact headers from "09:00, 10:00, 11:00 => 09:00-10:00, 10:00-11:00"
	// This shortens the list by one and shifts times left...
	slotHeaders := make([]string, len(timeSlots)-1)
	for i := 0; i < len(timeSlots)-1; i++ {
		slotHeaders[i] = timeSlots[i].Format("15:04") + "-" + timeSlots[i+1].Format("15:04")
	}

	dateStr := df.Col("date").Elem(0).String()
	date, err := time.Parse(time.DateOnly, dateStr)
	if err != nil {
		return DaySchedule{}, errors.New("Error parsing day date: " + err.Error())
	}
	daySchedule := DaySchedule{
		dateStr: dateStr,
		dayStr:  date.Weekday().String(),
		weekStr: df.Col("weekNumber").Elem(0).String(),
		shifts:  shiftRows,
		headers: append([]string{"Arbetstid", "Namn", "Tele"}, slotHeaders...),
	}

	return daySchedule, nil
}

/*
================================================================================
Handle footer from file
================================================================================
*/

// PrepareFooter extracts data+styles once and returns a reusable slice
func PrepareFooter(r io.Reader, sheet string) ([]FooterCell, error) {
	var footer []FooterCell

	srcFile, err := excelize.OpenReader(r)
	if err != nil {
		return footer, errors.New("Error opening footer file: " + err.Error())
	}

	rows, err := srcFile.GetRows(sheet)
	if err != nil {
		return nil, err
	}

	for rowIdx, row := range rows {
		for colIdx, val := range row {
			colNum := colIdx + 1
			rowNum := rowIdx + 1
			cellRef, _ := excelize.CoordinatesToCellName(colNum, rowNum)

			styleID, _ := srcFile.GetCellStyle(sheet, cellRef)
			styleJSON, _ := srcFile.GetStyle(styleID) // get full style definition

			footer = append(footer, FooterCell{
				Row:   rowNum,
				Col:   colNum,
				Value: val,
				Style: styleJSON,
			})
		}
	}

	return footer, nil
}

// ApplyFooterToSheet inserts the prepared footer into a new file
func ApplyFooterToSheet(dstFile *excelize.File, dstSheet string, footer []FooterCell, rowOffset int) error {
	for _, cell := range footer {
		cellName, _ := excelize.CoordinatesToCellName(cell.Col, cell.Row+rowOffset)
		if err := dstFile.SetCellValue(dstSheet, cellName, cell.Value); err != nil {
			return err
		}

		if cell.Style != nil {
			styleID, err := dstFile.NewStyle(cell.Style)
			if err != nil {
				return err
			}
			if err := dstFile.SetCellStyle(dstSheet, cellName, cellName, styleID); err != nil {
				return err
			}
		}
	}

	return nil
}

/*
================================================================================
Only Styles below this line
================================================================================
*/
func setStyles(f *excelize.File) error {
	var err error
	// TITLE ROW
	styleTitle, err = f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#AAAAAA"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Font: &excelize.Font{
			Bold:  true,
			Size:  20,
			Color: "FFFFFF",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	if err != nil {
		return errors.New("Failed to create styleTitle: " + err.Error())
	}
	// HEADER ROW
	styleHeader, err = f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#A1C2F1"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Font: &excelize.Font{
			Bold: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	if err != nil {
		return errors.New("Failed to create styleName: " + err.Error())
	}
	// NAME COLUMN
	styleName, err = f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#F4D793"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return errors.New("Failed to create styleName: " + err.Error())
	}
	// HOUR CELLS (depending on activity)
	styleFree, err = f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#B4B4B8"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	if err != nil {
		return errors.New("Failed to create styleFree: " + err.Error())
	}
	styleWork, err = f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#FFFFFF"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	if err != nil {
		return errors.New("Failed to create styleWork: " + err.Error())
	}
	styleLunch, err = f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#F6EFBD"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	if err != nil {
		return errors.New("Failed to create styleLunch: " + err.Error())
	}
	styleAssigned, err = f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#FFFFFF"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	if err != nil {
		return errors.New("Failed to create styleAssigned: " + err.Error())
	}

	return nil
}
