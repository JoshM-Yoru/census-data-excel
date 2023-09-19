package excel

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func updater() {
	file, err := excelize.OpenFile("Results.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// file_to_update, err := excelize.OpenFile(("GeocodeResults_Main_File.xlsx"))
	if err != nil {
		log.Fatal(err)
	}

    number_of_rows, err := file.GetRows("Sheet1")

    fmt.Println(number_of_rows);
}
