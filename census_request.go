package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
)

func census_request() {

	// url := CENSUS_API_URL + "/geocoder/geographies/addressbatch"
    csv := len(check_csv())
    fmt.Println(csv)

}

func check_csv() [][]string {

	file_name := "data.csv"

	file, err := os.Open(file_name)
	if err != nil {
		log.Fatal(err)
	}

	csv_Reader := csv.NewReader(file)
	csv_data, err := csv_Reader.ReadAll()
	if err != nil {
		log.Fatal(err)
	}

	csv_rows := len(csv_data)

	if csv_rows > 10000 {
		log.Fatal("Too Many Rows in the file '" + file_name + "'!")
	}

    return csv_data
}
