package main

import (
	"encoding/csv"
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

func csvConverter() {
	xlsx := excelize.NewFile()

	csvFileName := "test.csv"
	csvfile, err := os.Open(csvFileName)
	if err != nil {
		fmt.Println("Error opening CSV file:", err)
		return
	}
	defer csvfile.Close()

	csvReader, err := csv.NewReader(csvfile).ReadAll()
	sheetName := "Sheet1"

	for rowIndex, row := range csvReader {
		for colIndex, cellValue := range row {
			cellName, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
			xlsx.SetCellValue(sheetName, cellName, cellValue)
		}
	}

	xlsxFileName := "output.xlsx"
	if err := xlsx.SaveAs(xlsxFileName); err != nil {
		fmt.Println("Error saving XLSX file:", err)
		return
	}

	fmt.Println("CSV to XLSX conversion successful.")
}
