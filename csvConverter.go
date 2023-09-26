package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func csvConverter() {

    fmt.Println("Coverting csv result from Census.gov to an Excel file...")

	xlsx := excelize.NewFile()

	csvFileName := "Results.csv"
	csvfile, err := os.Open(csvFileName)
	if err != nil {
		fmt.Println("Error opening CSV file:", err)
		return
	}
	defer csvfile.Close()

	csvReader := csv.NewReader(csvfile)
	csvReader.FieldsPerRecord = -1
	csvMatrix, err := csvReader.ReadAll()
	if err != nil {
		fmt.Println("Something went wrong! ", err)
		os.Exit(0)
	}
	sheetName := "Sheet1"

	for rowIndex, row := range csvMatrix {
		for colIndex, cellValue := range row {
			if colIndex == 0 || colIndex == 8 || colIndex == 9 || colIndex == 10 || colIndex == 11 {
				intCell, err := strconv.Atoi(cellValue)
				if err != nil {
				}
				cellName, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
				xlsx.SetCellValue(sheetName, cellName, intCell)
			} else {
				cellName, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
				xlsx.SetCellValue(sheetName, cellName, cellValue)
			}
		}
	}

	xlsxFileName := "Results.xlsx"
	if err := xlsx.SaveAs(xlsxFileName); err != nil {
		fmt.Println("Error saving XLSX file:", err)
		return
	}

	fmt.Println("CSV to XLSX conversion successful.")
}
