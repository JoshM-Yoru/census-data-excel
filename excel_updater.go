package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"strconv"
	"sync"

	"github.com/schollz/progressbar/v3"
	"github.com/xuri/excelize/v2"
)
var mainMap map[int]int

func updater() {

	fmt.Println("Updating Main Excel table...")

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

	mainFile, err := excelize.OpenFile("GeocodeResults_Main_File - Copy(3).xlsx")
	if err != nil {
		log.Fatal(err)
	}

	sheet := "Sheet1"

	numberOfRows := len(csvMatrix)

	mainMap = createMainFileMap(mainFile)

	bar := progressbar.Default(int64(numberOfRows))
	defer bar.Finish()

	var wg sync.WaitGroup

	resultsCh := make(chan struct {
		id    string
		index int
	}, numberOfRows)

	for i := 0; i < numberOfRows; i++ {
		wg.Add(1)

		go processRow(&wg, &csvMatrix, sheet, i, resultsCh)
	}

	go func() {
		wg.Wait()
		close(resultsCh)
	}()

	go findMatches(mainFile, &csvMatrix, resultsCh)

	for i := 1; i <= numberOfRows; i++ {
		bar.Add(1)
	}

	mainFile.Save()
	fmt.Println("Updating Complete!")
}

func processRow(wg *sync.WaitGroup, file *[][]string, sheet string, rowIndex int, resultsCh chan<- struct {
	id    string
	index int
}) {
	defer wg.Done()

    contents := *file

	results_cell := contents[rowIndex][2]

	if results_cell == "Match" {
		id := contents[rowIndex][0]

		resultsCh <- struct {
			id    string
			index int
		}{id: id, index: rowIndex}
	}
}

func findMatches(mainFile *excelize.File, file *[][]string, resultsCh <-chan struct {
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
			updateRow(mainFile, file, row, result.index)
		} else {
			numberOfRows++
			updateRow(mainFile, file, numberOfRows, result.index)
		}
	}

	mainFile.Save()
}

func updateRow(mainFile *excelize.File, file *[][]string, mainRowIndex, rowIndex int) {

	mainSheet := "GeocodeResults (2)"
    contents := *file

    id := contents[rowIndex][0]
    originalAddress := contents[rowIndex][1]
    match := contents[rowIndex][2]
    exact := contents[rowIndex][3]
    address := contents[rowIndex][4]
    coordinates := contents[rowIndex][5]
    unk := contents[rowIndex][6]
    side := contents[rowIndex][7]
    state_id := contents[rowIndex][8]
    county := contents[rowIndex][9]
    group := contents[rowIndex][10]
    block := contents[rowIndex][11]

	id_int, err := strconv.Atoi(id)
    if err != nil {
        fmt.Println("Not a number:", id)
        return
    }
	state_id_int, err := strconv.Atoi(state_id)
    if err != nil {
        fmt.Println("Not a number:", id)
        return
    }
	county_int, err := strconv.Atoi(county)
    if err != nil {
        fmt.Println("Not a number:", id)
        return
    }
	group_int, err := strconv.Atoi(group)
    if err != nil {
        fmt.Println("Not a number:", id)
        return
    }
	block_int, err := strconv.Atoi(block)
    if err != nil {
        fmt.Println("Not a number:", id)
        return
    }

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

func createMainFileMap(file *excelize.File) map[int]int {
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

