package main

import (
	"fmt"

	"io"

	"github.com/rsravanreddy/go-office-parser/parser"
)

func main() {

	var dr io.Reader
	dr, _ = parser.NewDocxReader("./testdata/demo.docx")

	buf := make([]byte, 4000)
	s := ""
	for {
		n, err := dr.Read(buf)
		s += string(buf[:n])
		if err == io.EOF {
			break
		}
	}
	fmt.Println(s)

	dr, _ = parser.NewXlsxReader("./testdata/sample.xlsx")

	buf = make([]byte, 4000)
	s = ""
	for {
		n, err := dr.Read(buf)
		s += string(buf[:n])
		if err == io.EOF {
			break
		}
	}
	fmt.Println(s)

	// res, err := xlsxparser.Parse("./testdata/sample.xlsx")
	// elapsed := time.Since(start)
	// log.Printf("parsing took %s", elapsed)
	// if err == nil {
	// 	fmt.Println(res)
	// } else {
	// 	fmt.Println(err)

	// }

}
