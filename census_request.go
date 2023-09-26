package main

import (
	"fmt"
	"os/exec"

	"log"
	"os"
)

func curlRequest() {
    curl := "curl"
	curlArgs := []string{"--form", "addressFile=@data.csv", "--form", "benchmark=Public_AR_Current", "--form", "vintage=Current_Current", "https://geocoding.geo.census.gov/geocoder/geographies/addressbatch", "--output", "Results.csv"}

    cmd := exec.Command(curl, curlArgs...)

    cmd.Stdout = os.Stdout
    cmd.Stderr = os.Stderr

	err := cmd.Run()
	if err != nil {
	    log.Fatal(err)
	}

	fmt.Println("Curl command executed successfully!")

}
