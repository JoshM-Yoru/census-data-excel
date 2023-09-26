package main

import (
	"fmt"
	"log"
	"strconv"
	"time"

	"github.com/schollz/progressbar/v3"
	"github.com/xuri/excelize/v2"
)

func match_finder() {

	fmt.Println("Updating Main Excel table...")

	file, err := excelize.OpenFile("Results.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	mainFile, err := excelize.OpenFile("GeocodeResults_Main_File.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	sheet := "Sheet1"

	rows, err := file.GetRows(sheet)
	if err != nil {
		log.Fatal(err)
	}

	number_of_rows := len(rows)

	bar := progressbar.Default(int64(number_of_rows))

	for i := 1; i <= number_of_rows; i++ {
		results_cell, err := file.GetCellValue(sheet, "C"+strconv.Itoa(i))
		if err != nil {
			log.Fatal(err)
		}

		bar.Add(1)
		time.Sleep(40 * time.Millisecond)

		if results_cell == "Match" {
			id, err := file.GetCellValue(sheet, "A"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			originalAddress, err := file.GetCellValue(sheet, "B"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			match, err := file.GetCellValue(sheet, "C"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			exact, err := file.GetCellValue(sheet, "D"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			address, err := file.GetCellValue(sheet, "E"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			coordinates, err := file.GetCellValue(sheet, "F"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			unk, err := file.GetCellValue(sheet, "G"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			side, err := file.GetCellValue(sheet, "H"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			state_id, err := file.GetCellValue(sheet, "I"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			county, err := file.GetCellValue(sheet, "J"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			group, err := file.GetCellValue(sheet, "K"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			block, err := file.GetCellValue(sheet, "L"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			excel_updater(mainFile, id, originalAddress, match, exact, address, coordinates, unk, side, state_id, county, group, block)
		}
	}

	mainFile.Save()
	fmt.Println("Updating Complete!")
}

func excel_updater(file *excelize.File, id string, originalAddress string, match string, exact string, address string, coordinates string, unk string, side string, state_id string, county string, group string, block string) {

	sheet := "GeocodeResults (2)"

	rows, err := file.GetRows(sheet)
	if err != nil {
		log.Fatal(err)
	}

	number_of_rows := len(rows)

	id_int, err := strconv.Atoi(id)
	state_id_int, err := strconv.Atoi(state_id)
	county_int, err := strconv.Atoi(county)
	group_int, err := strconv.Atoi(group)
	block_int, err := strconv.Atoi(block)

	found := false

	for i := 1; i <= number_of_rows; i++ {
		results_cell, err := file.GetCellValue(sheet, "A"+strconv.Itoa(i))
		if err != nil {
			log.Fatal("Invalid cell: ", err)
		}

		if results_cell == id {
			found = true
			file.SetCellStr(sheet, "C"+strconv.Itoa(i), match)
			file.SetCellStr(sheet, "D"+strconv.Itoa(i), exact)
			file.SetCellStr(sheet, "E"+strconv.Itoa(i), address)
			file.SetCellStr(sheet, "F"+strconv.Itoa(i), coordinates)
			file.SetCellStr(sheet, "G"+strconv.Itoa(i), unk)
			file.SetCellStr(sheet, "H"+strconv.Itoa(i), side)
			file.SetCellInt(sheet, "I"+strconv.Itoa(i), state_id_int)
			file.SetCellInt(sheet, "J"+strconv.Itoa(i), county_int)
			file.SetCellInt(sheet, "K"+strconv.Itoa(i), group_int)
			file.SetCellInt(sheet, "L"+strconv.Itoa(i), block_int)
		}
	}

	if !found {
		file.SetCellInt(sheet, "A"+strconv.Itoa(number_of_rows+1), id_int)
		file.SetCellStr(sheet, "B"+strconv.Itoa(number_of_rows+1), originalAddress)
		file.SetCellStr(sheet, "C"+strconv.Itoa(number_of_rows+1), match)
		file.SetCellStr(sheet, "D"+strconv.Itoa(number_of_rows+1), exact)
		file.SetCellStr(sheet, "E"+strconv.Itoa(number_of_rows+1), address)
		file.SetCellStr(sheet, "F"+strconv.Itoa(number_of_rows+1), coordinates)
		file.SetCellStr(sheet, "G"+strconv.Itoa(number_of_rows+1), unk)
		file.SetCellStr(sheet, "H"+strconv.Itoa(number_of_rows+1), side)
		file.SetCellInt(sheet, "I"+strconv.Itoa(number_of_rows+1), state_id_int)
		file.SetCellInt(sheet, "J"+strconv.Itoa(number_of_rows+1), county_int)
		file.SetCellInt(sheet, "K"+strconv.Itoa(number_of_rows+1), group_int)
		file.SetCellInt(sheet, "L"+strconv.Itoa(number_of_rows+1), block_int)
	}
}
