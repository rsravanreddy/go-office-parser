package main

import (
	"fmt"
	"log"
	"os"
	"time"

	"io"

	"github.com/richardlehane/mscfb"
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
		fmt.Println(len(s))

		dr.Close()

		dr, _ = parser.NewXlsxReader("./testdata/sensitive123.xlsx")

		buf = make([]byte, 4000)
		s = ""
		for {
			n, err := dr.Read(buf)
			s = string(buf[:n])
			if err == io.EOF {
				break
			}
		}
		fmt.Println(len(s))
		dr.Close()

		// res, err := xlsxparser.Parse("./testdata/sample.xlsx")
		// elapsed := time.Since(start)
		// log.Printf("parsing took %s", elapsed)
		// if err == nil {
		// 	fmt.Println(res)
		// } else {
		// 	fmt.Println(err)

		// }

		file, _ := os.Open("testdata/test.doc")
		defer file.Close()
		doc, err := mscfb.New(file)
		if err != nil {
			log.Fatal(err)
		}
		for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
			//buf := make([]byte, 1024)
			//_, _ := doc.Read(buf)
			// if i > 0 {
			// 	fmt.Println("data" + string(buf[:i]))
			// }
			fmt.Println(entry.Name)
		}

	}

	for {
		time.Sleep(100 * time.Millisecond)
		fmt.Println("sleeping..")

	}
}
