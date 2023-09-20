package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func excel_updater() {
	file, err := excelize.OpenFile("Results.xlsx")
	if err != nil {
        fmt.Println("wtf")
		log.Fatal(err)
	}

	// file_to_update, err := excelize.OpenFile(("GeocodeResults_Main_File.xlsx"))
	if err != nil {
		log.Fatal(err)
	}

    fmt.Println(file.GetCellValue("GeocodeResults (3)", "B2"))

    number_of_rows, err := file.GetRows("GeocodeResults (3)")

    fmt.Println(len(number_of_rows));
}
