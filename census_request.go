package main

import (
	"bytes"
	"fmt"
	"io"
	"mime/multipart"
	"net/http"
	"os"

	"github.com/schollz/progressbar/v3"
)

func httpFilePost() {
	fmt.Println("Querying Census.gov for address matches...")

	apiEndpoint := "https://geocoding.geo.census.gov/geocoder/geographies/addressbatch"

	file, err := os.Open("data.csv")
	if err != nil {
		fmt.Println("Error opening file:", err)
		return
	}
	defer file.Close()

	fileInfo, err := file.Stat()
	if err != nil {
		fmt.Println("Error getting file info:", err)
		return
	}
	fileSize := fileInfo.Size()

	body := &bytes.Buffer{}

	writer := multipart.NewWriter(body)

	part, err := writer.CreateFormFile("addressFile", "data.csv")
	if err != nil {
		fmt.Println("Error creating form file:", err)
		return
	}

	bar := progressbar.DefaultBytes(
		fileSize,
		"Uploading",
	)

	multiWriter := io.MultiWriter(part, bar)

	_, err = io.Copy(multiWriter, file)
	if err != nil {
		fmt.Println("Error copying file content:", err)
		return
	}

	writer.WriteField("benchmark", "Public_AR_Current")
	writer.WriteField("vintage", "Current_Current")

	writer.Close()

	req, err := http.NewRequest("POST", apiEndpoint, body)
	if err != nil {
		fmt.Println("Error creating request:", err)
		return
	}

	req.Header.Set("Content-Type", writer.FormDataContentType())

	fmt.Println("Waiting for a response...")

	client := http.Client{}
	res, err := client.Do(req)
	if err != nil {
		fmt.Println("Error making request:", err)
		return
	}
	defer res.Body.Close()

	if res.StatusCode == http.StatusOK {
		outputFile, err := os.Create("Results.csv")
		if err != nil {
			fmt.Println("Error creating output file:", err)
			return
		}
		defer outputFile.Close()

		responseSize := res.ContentLength
		downloadBar := progressbar.DefaultBytes(
			responseSize,
			"Downloading",
		)

		multiWriterDownload := io.MultiWriter(outputFile, downloadBar)

		_, err = io.Copy(multiWriterDownload, res.Body)
		if err != nil {
			fmt.Println("Error copying response to file:", err)
			return
		}

		fmt.Println("File uploaded successfully!")
	} else {
		fmt.Println("Error uploading file. Status code:", res.StatusCode)
	}
}

