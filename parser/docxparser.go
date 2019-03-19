package parser

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

type DocxReader struct {
	err    error
	data   []byte
	offset int
	length int
}

func NewDocxReader(path string) (*DocxReader, error) {
	dr := &DocxReader{}
	dr.offset = 0
	dr.length = 0
	var data string
	data, dr.err = dr.parse(path)
	dr.data = make([]byte, len(data))
	copy(dr.data[:], data[:])
	dr.length = len(dr.data)
	return dr, dr.err
}

func (r *DocxReader) Read(b []byte) (int, error) {

	if r.err != nil {
		return 0, r.err
	}
	if r.offset-r.length == 0 {
		return 0, io.EOF
	}
	len := util.Min(len(b), r.length-r.offset)
	copy(b[:], r.data[r.offset:])
	r.offset = r.offset + len
	return len, nil

}

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
func (dr *DocxReader) parse(path string) (data string, err error) {
	if !util.FileExists(path) {
		return data, errors.New("file does not exist")
	}
	r, err := zip.OpenReader(path)
	if err != nil || r == nil {
		return data, err
	}
	defer r.Close()
	if err != nil {
		return data, err
	}

	file := util.RetrieveWordDoc(r.File)

	if file == nil {
		return data, errors.New("file is not valid docx")
	}

	rc, err := file.Open()
	if err != nil {
		return data, err

	}
	ByteData, _ := util.ReadFile(rc)

	// read our opened xmlFile as a byte array.
	byteValue := []byte(ByteData)

	var doc document
	err = xml.Unmarshal(byteValue, &doc)
	if err != nil {
		log.Fatal(err)
	}

	var b body

	err = xml.Unmarshal([]byte(doc.Wp.InnerXML), &b)
	if err != nil {
		return data, errors.New("file is not valid docx")

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
				sa := []string{data, element.Value}
				data = strings.Join(sa, " ")
			}
		}

	}
	return data, err
}
