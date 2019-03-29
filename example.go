package main

import (
	"fmt"

	"io"

	"github.com/rsravanreddy/go-office-parser/parser"
)

func main() {

	{
		var dr io.ReadCloser
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

		dr.Close()

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
		dr.Close()

	}

}
