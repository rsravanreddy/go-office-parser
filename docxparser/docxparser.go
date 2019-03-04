package docxparser

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"log"
	"strings"

	"github.com/rsravanreddy/go-office-parser/util"
)

type document struct {
	Wp struct {
		InnerXML string `xml:",innerxml"`
	} `xml:"body"`
}

type doc struct {
	test string
}
type body struct {
	//XMLName xml.Name `xml:"w body"`
	Wp []record `xml:"p"`
}

type wp struct {
	XMLName xml.Name `xml:"p"`
	Records []record `xml:"r"`
}

type record struct {
	XMLName xml.Name `xml:"r"`
	Value   string   `xml:"t"`
}

// Parse ... parses a word document and returns a string
func Parse(path string) (result string, err error) {
	if !util.FileExists(path) {
		return result, errors.New("file does not exist")
	}
	r, err := zip.OpenReader(path)
	defer r.Close()
	if err != nil || r == nil {
		return result, err
	}
	if err != nil {
		return result, err
	}

	file := util.RetrieveWordDoc(r.File)

	if file == nil {
		return result, errors.New("file is not valid docx")
	}

	rc, err := file.Open()
	if err != nil {
		return result, err

	}
	data, _ := util.ReadFile(rc)

	// read our opened xmlFile as a byte array.
	byteValue := []byte(data)

	var doc document
	err = xml.Unmarshal(byteValue, &doc)
	if err != nil {
		log.Fatal(err)
	}

	var b body

	err = xml.Unmarshal([]byte(doc.Wp.InnerXML), &b)
	if err != nil {
		return result, errors.New("file is not valid docx")

	}

	d := xml.NewDecoder(bytes.NewBufferString(doc.Wp.InnerXML))
	for {
		var t wp
		err := d.Decode(&t)
		if err == io.EOF {
			break
		}
		if err == nil {
			for _, element := range t.Records {
				sa := []string{result, element.Value}
				result = strings.Join(sa, " ")
			}
		}

	}
	return result, err
}
