package main

import (
	"fmt"
	"log"
	"strconv"
	"sync"
	"time"

	"github.com/schollz/progressbar/v3"
	"github.com/xuri/excelize/v2"
)

// func match_finder() {
//
// 	fmt.Println("Updating Main Excel table...")
//
// 	file, err := excelize.OpenFile("Results.xlsx")
// 	if err != nil {
// 		log.Fatal(err)
// 	}
//
// 	mainFile, err := excelize.OpenFile("GeocodeResults_Main_File.xlsx")
// 	if err != nil {
// 		log.Fatal(err)
// 	}
//
// 	sheet := "Sheet1"
//
// 	rows, err := file.GetRows(sheet)
// 	if err != nil {
// 		log.Fatal(err)
// 	}
//
// 	numberOfRows := len(rows)
//
//     mainMap := mapConverter(mainFile)
//
// 	bar := progressbar.Default(int64(numberOfRows))
//     defer bar.Finish()
//
// 	for i := 1; i <= numberOfRows; i++ {
// 		results_cell, err := file.GetCellValue(sheet, "C"+strconv.Itoa(i))
// 		if err != nil {
// 			log.Fatal(err)
// 		}
//
// 		bar.Add(1)
// 		time.Sleep(40 * time.Millisecond)
//
// 		if results_cell == "Match" {
// 			id, err := file.GetCellValue(sheet, "A"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			originalAddress, err := file.GetCellValue(sheet, "B"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			match, err := file.GetCellValue(sheet, "C"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			exact, err := file.GetCellValue(sheet, "D"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			address, err := file.GetCellValue(sheet, "E"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			coordinates, err := file.GetCellValue(sheet, "F"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			unk, err := file.GetCellValue(sheet, "G"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			side, err := file.GetCellValue(sheet, "H"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			state_id, err := file.GetCellValue(sheet, "I"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			county, err := file.GetCellValue(sheet, "J"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			group, err := file.GetCellValue(sheet, "K"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			block, err := file.GetCellValue(sheet, "L"+strconv.Itoa(i))
// 			if err != nil {
// 				log.Fatal(err)
// 			}
//
// 			excel_updater_map(mainFile, mainMap, id, originalAddress, match, exact, address, coordinates, unk, side, state_id, county, group, block)
// 		}
// 	}
//
// 	mainFile.Save()
// 	fmt.Println("Updating Complete!")
// }
//
// func excel_updater_map(file *excelize.File, mapToCheck map[int]int, id string, originalAddress string, match string, exact string, address string, coordinates string, unk string, side string, state_id string, county string, group string, block string) {
//
// 	sheet := "GeocodeResults (2)"
//
// 	rows, err := file.GetRows(sheet)
// 	if err != nil {
// 		log.Fatal(err)
// 	}
//
// 	numberOfRows := len(rows)
//
// 	id_int, err := strconv.Atoi(id)
// 	state_id_int, err := strconv.Atoi(state_id)
// 	county_int, err := strconv.Atoi(county)
// 	group_int, err := strconv.Atoi(group)
// 	block_int, err := strconv.Atoi(block)
//
// 	if row, ok := mapToCheck[id_int]; ok {
// 		file.SetCellStr(sheet, "C"+strconv.Itoa(row), match)
// 		file.SetCellStr(sheet, "D"+strconv.Itoa(row), exact)
// 		file.SetCellStr(sheet, "E"+strconv.Itoa(row), address)
// 		file.SetCellStr(sheet, "F"+strconv.Itoa(row), coordinates)
// 		file.SetCellStr(sheet, "G"+strconv.Itoa(row), unk)
// 		file.SetCellStr(sheet, "H"+strconv.Itoa(row), side)
// 		file.SetCellInt(sheet, "I"+strconv.Itoa(row), state_id_int)
// 		file.SetCellInt(sheet, "J"+strconv.Itoa(row), county_int)
// 		file.SetCellInt(sheet, "K"+strconv.Itoa(row), group_int)
// 		file.SetCellInt(sheet, "L"+strconv.Itoa(row), block_int)
// 	} else {
// 		file.SetCellInt(sheet, "A"+strconv.Itoa(numberOfRows+1), id_int)
// 		file.SetCellStr(sheet, "B"+strconv.Itoa(numberOfRows+1), originalAddress)
// 		file.SetCellStr(sheet, "C"+strconv.Itoa(numberOfRows+1), match)
// 		file.SetCellStr(sheet, "D"+strconv.Itoa(numberOfRows+1), exact)
// 		file.SetCellStr(sheet, "E"+strconv.Itoa(numberOfRows+1), address)
// 		file.SetCellStr(sheet, "F"+strconv.Itoa(numberOfRows+1), coordinates)
// 		file.SetCellStr(sheet, "G"+strconv.Itoa(numberOfRows+1), unk)
// 		file.SetCellStr(sheet, "H"+strconv.Itoa(numberOfRows+1), side)
// 		file.SetCellInt(sheet, "I"+strconv.Itoa(numberOfRows+1), state_id_int)
// 		file.SetCellInt(sheet, "J"+strconv.Itoa(numberOfRows+1), county_int)
// 		file.SetCellInt(sheet, "K"+strconv.Itoa(numberOfRows+1), group_int)
// 		file.SetCellInt(sheet, "L"+strconv.Itoa(numberOfRows+1), block_int)
// 	}
// }
//
// func mapConverter(file *excelize.File) map[int]int {
// 	sheet := "GeocodeResults (2)"
//
// 	rows, err := file.GetRows(sheet)
// 	if err != nil {
// 		log.Fatal(err)
// 	}
//
// 	numberOfRows := len(rows)
//
// 	mainMap := make(map[int]int)
//
// 	for i := 1; i <= numberOfRows; i++ {
// 		results_cell, err := file.GetCellValue(sheet, "A"+strconv.Itoa(i))
// 		if err != nil {
// 			log.Fatal("Invalid cell: ", err)
// 		}
// 		if results_cell != "PK_Id" {
// 			key, err := strconv.Atoi(results_cell)
// 			if err != nil {
// 				log.Fatal("Not a string: ", err)
// 			}
// 			mainMap[key] = i
// 		}
// 	}
//
// 	return mainMap
// }

var mainMap map[int]int

func matchFinderConcurrent() {

	fmt.Println("Updating Main Excel table...")

	file, err := excelize.OpenFile("Results.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	mainFile, err := excelize.OpenFile("GeocodeResults_Main_File - Copy(1).xlsx")
	if err != nil {
		log.Fatal(err)
	}

	sheet := "Sheet1"

	rows, err := file.GetRows(sheet)
	if err != nil {
		log.Fatal(err)
	}

	numberOfRows := len(rows)

	mainMap = mapConverter(mainFile)

	bar := progressbar.Default(int64(numberOfRows))
	defer bar.Finish()

	var wg sync.WaitGroup

	resultsCh := make(chan struct {
		id    string
		index int
	}, numberOfRows)

	for i := 1; i <= numberOfRows; i++ {
		wg.Add(1)

		go procssRow(&wg, file, sheet, i, resultsCh)
	}

	go func() {
		wg.Wait()
		close(resultsCh)
	}()

	go updateMainFile(mainFile, file, resultsCh)

	for i := 1; i <= numberOfRows; i++ {
		bar.Add(1)
		time.Sleep(40 * time.Millisecond)
	}

	mainFile.Save()
	fmt.Println("Updating Complete!")
}

func procssRow(wg *sync.WaitGroup, file *excelize.File, sheet string, rowIndex int, resultsCh chan<- struct {
	id    string
	index int
}) {
	defer wg.Done()

	results_cell, err := file.GetCellValue(sheet, "C"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	if results_cell == "Match" {
		id, err := file.GetCellValue(sheet, "A"+strconv.Itoa(rowIndex))
		if err != nil {
			log.Fatal(err)
		}

		resultsCh <- struct {
			id    string
			index int
		}{id: id, index: rowIndex}
	}
}

func updateMainFile(mainFile *excelize.File, file *excelize.File, resultsCh <-chan struct {
	id    string
	index int
}) {
	mainSheet := "GeocodeResults (2)"

	rows, err := mainFile.GetRows(mainSheet)
	if err != nil {
		fmt.Println("Unable to get number of rows:", err)
		return
	}

	numberOfRows := len(rows)

	for result := range resultsCh {
		id, err := strconv.Atoi(result.id)
		if err != nil {
			fmt.Println("Unable to convert to int:", err)
			return
		}
		row, ok := mainMap[id]
		if ok {
			excelUpdaterMap(mainFile, file, row, result.index)
		} else {
			numberOfRows++
			excelUpdaterMap(mainFile, file, numberOfRows, result.index)
		}
	}

	mainFile.Save()
}

func excelUpdaterMap(mainFile *excelize.File, file *excelize.File, mainRowIndex, rowIndex int) {

    sheet := "Sheet1"
	mainSheet := "GeocodeResults (2)"

	id, err := file.GetCellValue(sheet, "A"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	originalAddress, err := file.GetCellValue(sheet, "B"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	match, err := file.GetCellValue(sheet, "C"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	exact, err := file.GetCellValue(sheet, "D"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	address, err := file.GetCellValue(sheet, "E"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	coordinates, err := file.GetCellValue(sheet, "F"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	unk, err := file.GetCellValue(sheet, "G"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	side, err := file.GetCellValue(sheet, "H"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	state_id, err := file.GetCellValue(sheet, "I"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	county, err := file.GetCellValue(sheet, "J"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	group, err := file.GetCellValue(sheet, "K"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	block, err := file.GetCellValue(sheet, "L"+strconv.Itoa(rowIndex))
	if err != nil {
		log.Fatal(err)
	}

	id_int, err := strconv.Atoi(id)
	state_id_int, err := strconv.Atoi(state_id)
	county_int, err := strconv.Atoi(county)
	group_int, err := strconv.Atoi(group)
	block_int, err := strconv.Atoi(block)

	mainFile.SetCellInt(mainSheet, "A"+strconv.Itoa(mainRowIndex), id_int)
	mainFile.SetCellStr(mainSheet, "B"+strconv.Itoa(mainRowIndex), originalAddress)
	mainFile.SetCellStr(mainSheet, "C"+strconv.Itoa(mainRowIndex), match)
	mainFile.SetCellStr(mainSheet, "D"+strconv.Itoa(mainRowIndex), exact)
	mainFile.SetCellStr(mainSheet, "E"+strconv.Itoa(mainRowIndex), address)
	mainFile.SetCellStr(mainSheet, "F"+strconv.Itoa(mainRowIndex), coordinates)
	mainFile.SetCellStr(mainSheet, "G"+strconv.Itoa(mainRowIndex), unk)
	mainFile.SetCellStr(mainSheet, "H"+strconv.Itoa(mainRowIndex), side)
	mainFile.SetCellInt(mainSheet, "I"+strconv.Itoa(mainRowIndex), state_id_int)
	mainFile.SetCellInt(mainSheet, "J"+strconv.Itoa(mainRowIndex), county_int)
	mainFile.SetCellInt(mainSheet, "K"+strconv.Itoa(mainRowIndex), group_int)
	mainFile.SetCellInt(mainSheet, "L"+strconv.Itoa(mainRowIndex), block_int)
}

func mapConverter(file *excelize.File) map[int]int {
	sheet := "GeocodeResults (2)"

	rows, err := file.GetRows(sheet)
	if err != nil {
		log.Fatal(err)
	}

	numberOfRows := len(rows)

	mainMap := make(map[int]int)

	for i := 1; i <= numberOfRows; i++ {
		results_cell, err := file.GetCellValue(sheet, "A"+strconv.Itoa(i))
		if err != nil {
			log.Fatal("Invalid cell: ", err)
		}
		if results_cell != "PK_Id" {
			key, err := strconv.Atoi(results_cell)
			if err != nil {
				log.Fatal("Not a string: ", err)
			}
			mainMap[key] = i
		}
	}

	return mainMap
}
