package main

import (
	"docxparser"
	"fmt"
	"log"
	"time"
	"xlsxparser"
)

func main() {
	start := time.Now()

	res, err := docxparser.Parse("./testdata/demo.docx")

	if err == nil {
		fmt.Println(res)
	} else {
		fmt.Println(err)

	}

	res, err = xlsxparser.Parse("./testdata/sample.xlsx")
	elapsed := time.Since(start)
	log.Printf("parsing took %s", elapsed)
	if err == nil {
		fmt.Println(res)
	} else {
		fmt.Println(err)

	}

}
