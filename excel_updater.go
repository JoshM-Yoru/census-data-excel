package main

import (
	"fmt"
	"log"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func match_finder() {
	file, err := excelize.OpenFile("Results.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	rows, err := file.GetRows("GeocodeResults (3)")
	if err != nil {
		log.Fatal(err)
	}

	number_of_rows := len(rows)

	for i := 1; i <= number_of_rows; i++ {
		results_cell, err := file.GetCellValue("GeocodeResults (3)", "C"+strconv.Itoa(i))
		if err != nil {
			log.Fatal(err)
		}

		if results_cell == "Match" {
			id, err := file.GetCellValue("GeocodeResults (3)", "A"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			match, err := file.GetCellValue("GeocodeResults (3)", "C"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			exact, err := file.GetCellValue("GeocodeResults (3)", "D"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			address, err := file.GetCellValue("GeocodeResults (3)", "E"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			coordinates, err := file.GetCellValue("GeocodeResults (3)", "F"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			unk, err := file.GetCellValue("GeocodeResults (3)", "G"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			side, err := file.GetCellValue("GeocodeResults (3)", "H"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			state_id, err := file.GetCellValue("GeocodeResults (3)", "I"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			county, err := file.GetCellValue("GeocodeResults (3)", "J"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			group, err := file.GetCellValue("GeocodeResults (3)", "K"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			block, err := file.GetCellValue("GeocodeResults (3)", "L"+strconv.Itoa(i))
			if err != nil {
				log.Fatal(err)
			}

			excel_updater(id, match, exact, address, coordinates, unk, side, state_id, county, group, block)
		}
	}

}

func excel_updater(id string, match string, exact string, address string, coordinates string, unk string, side string, state_id string, county string, group string, block string) {

	file, err := excelize.OpenFile("GeocodeResults_Main_File.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	rows, err := file.GetRows("GeocodeResults (2)")
	if err != nil {
		log.Fatal(err)
	}

	number_of_rows := len(rows)

	state_id_int, err := strconv.Atoi(state_id)
    county_int, err := strconv.Atoi(county)
    group_int, err := strconv.Atoi(group)
    block_int, err := strconv.Atoi(block)

	for i := 1; i <= number_of_rows; i++ {
		results_cell, err := file.GetCellValue("GeocodeResults (2)", "A"+strconv.Itoa(i))
		if err != nil {
			log.Fatal(err)
		}

		if results_cell == id {
			file.SetCellStr("GeocodeResults (2)", "C"+strconv.Itoa(i), match)
			file.SetCellStr("GeocodeResults (2)", "D"+strconv.Itoa(i), exact)
			file.SetCellStr("GeocodeResults (2)", "E"+strconv.Itoa(i), address)
			file.SetCellStr("GeocodeResults (2)", "F"+strconv.Itoa(i), coordinates)
			file.SetCellStr("GeocodeResults (2)", "G"+strconv.Itoa(i), unk)
			file.SetCellStr("GeocodeResults (2)", "H"+strconv.Itoa(i), side)
			file.SetCellInt("GeocodeResults (2)", "I"+strconv.Itoa(i), state_id_int)
			file.SetCellInt("GeocodeResults (2)", "J"+strconv.Itoa(i), county_int)
			file.SetCellInt("GeocodeResults (2)", "K"+strconv.Itoa(i), group_int)
			file.SetCellInt("GeocodeResults (2)", "L"+strconv.Itoa(i), block_int)
		}
	}

	file.Save()
}
