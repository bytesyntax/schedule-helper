package main

import "testing"

func BenchmarkCreateWeekSchedules(b *testing.B) {
	df, err := readAndRefineInputData("DailyStaffingSchedule_1348853_1494f5a70b167da50a8.xlsx")
	if err != nil {
		b.Error("Error reading input data:", err)
		return
	}

	for b.Loop() {
		createWeekSchedules(df)
	}
}
