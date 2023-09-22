package main

import (
	"bytes"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"strconv"

	"log"
	"os"
)

type api_params struct {
	benchmarks string `json:"benchmarks"`
	vintage    string `json:"vintage"`
	format     string `json:"format"`
}

type api_body struct {
	files  [][]string `json:"files"`
	params api_params `json:"params"`
}

func census_request() {

	url := CENSUS_API_URL + "/geocoder/geographies/addressbatch"
	data := check_csv()

	params := api_params{
		benchmarks: BENCHMARKS,
		vintage:    VINTAGES,
		format:     "json",
	}

	body := api_body{
		files:  data,
		params: params,
	}

	fmt.Println(body.params)

	requestBody, err := json.Marshal(body)
	if err != nil {
		log.Fatalf("Unable to marshal to json object: %v", err)
	}

	req, err := http.NewRequest("POST", url, bytes.NewBuffer(requestBody))
	if err != nil {
		log.Fatalf("Unable to reach server: %v", err)
	}
    
	code := "Status code: " + strconv.Itoa(req.Response.StatusCode) 
	defer fmt.Println(code)

	// defer req.Body.Close()

	responseBody, err := io.ReadAll(req.Body)
	if err != nil {
		log.Fatalf("No Response: %v", err)
	}

	response := "Response: " + string(responseBody)

	println(response)

}

func check_csv() [][]string {

	file_name := "data.csv"

	file, err := os.Open(file_name)
	if err != nil {
		log.Fatalf("Unable to open file: %v", err)
	}

	csv_Reader := csv.NewReader(file)
	csv_data, err := csv_Reader.ReadAll()
	if err != nil {
		log.Fatalf("Unable to read file: %v", err)
	}

	csv_rows := len(csv_data)

	println(csv_rows)

	if csv_rows > 10000 {
		log.Fatal("Too Many Rows in the file '" + file_name + "'!")
	}

	return csv_data
}
