package main

import (
	"fmt"
	"io"

	"github.com/rsravanreddy/go-office-parser/parser"
)

func main() {
	// defer profile.Start(profile.MemProfile).Stop()

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
	var totallen int
	buf = make([]byte, 4000)
	s = ""
	for {
		n, err := dr.Read(buf)
		totallen += n
		s += string(buf[:n])
		if err == io.EOF {
			break
		}

	}
	fmt.Println(s)
	dr.Close()

	drr := parser.NewDocReader("./testdata/test.doc")
	s = ""
	for {
		n, err := drr.Read(buf)
		s += string(buf[:n])
		if err == io.EOF {
			break
		}
	}
	drr.Close()
	println(s)

	xr := parser.NewXlsReader("./testdata/test.xls")
	s = ""
	for {
		n, err := xr.Read(buf)
		s += string(buf[:n])
		if err == io.EOF {
			break
		}
	}
	println(s)

}
