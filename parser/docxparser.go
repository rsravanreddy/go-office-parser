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
	err      error
	data     []byte
	offset   int
	xmlIndex int
	decoder  *xml.Decoder
}

func NewDocxReader(path string) (*DocxReader, error) {
	dr := &DocxReader{}
	dr.offset = 0
	dr.xmlIndex = 0
	dr.err = dr.parse(path)
	return dr, dr.err
}

func (r *DocxReader) Read(b []byte) (int, error) {

	if r.err != nil {
		return 0, r.err
	}
	//need to fill
	var err error
	var lenRead int
	if r.offset+len(b) > len(r.data) {
		lenRead, err = r.FillFromXml(r.offset + len(b) - len(r.data))
	}
	if err != nil && lenRead == 0 {
		return 0, err
	}
	len := util.Min(len(b), len(r.data)-r.offset)
	copy(b[:], r.data[r.offset:])
	//r.offset = r.offset + len
	r.data = r.data[len:]
	return len, nil

}

func (r *DocxReader) Close() error {
	r.err = errors.New("reader already closed")
	return nil
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
func (dr *DocxReader) parse(path string) (err error) {
	if !util.FileExists(path) {
		return errors.New("file does not exist")
	}
	r, err := zip.OpenReader(path)
	if err != nil || r == nil {
		return err
	}
	defer r.Close()
	if err != nil {
		return err
	}

	file := util.RetrieveWordDoc(r.File)

	if file == nil {
		return errors.New("file is not valid docx")
	}

	rc, err := file.Open()
	if rc != nil {
		defer rc.Close()
	}
	if err != nil {
		return err

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
		return errors.New("file is not valid docx")

	}

	dr.decoder = xml.NewDecoder(bytes.NewBufferString(doc.Wp.InnerXML))
	return err
}

func (dr *DocxReader) FillFromXml(minSize int) (int, error) {
	var err error
	var data string
	var dataSize int
	for {

		if dataSize > minSize {
			break
		}
		var t wp
		err = dr.decoder.Decode(&t)
		if err == io.EOF {
			break
		}
		if err == nil {
			for _, element := range t.Records {
				sa := []string{data, element.Value}
				dataSize += len([]byte(element.Value)) + 1
				data = strings.Join(sa, " ")
			}
		}
	}
	byteData := []byte(data)
	dr.data = append(dr.data, byteData...)
	dataSize += len(byteData)
	return dataSize, err
}
