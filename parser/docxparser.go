package parser

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"io"
	"strings"

	"github.com/rsravanreddy/go-office-parser/util"
)

type DocxReader struct {
	err           error
	data          []byte
	offset        int
	xmlIndex      int
	decoder       *xml.Decoder
	zipReadCloser io.ReadCloser
	workbookbRc   io.ReadCloser
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
	if r.zipReadCloser != nil {
		r.zipReadCloser.Close()
	}
	if r.workbookbRc != nil {
		r.workbookbRc.Close()
	}
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
	zipReadCloser, err := zip.OpenReader(path)
	if err != nil || zipReadCloser == nil {
		return err
	}
	if err != nil {
		return err
	}

	file := util.RetrieveWordDoc(zipReadCloser.File)

	if file == nil {
		return errors.New("file is not valid docx")
	}

	dr.workbookbRc, err = file.Open()

	dr.decoder = xml.NewDecoder(dr.workbookbRc)

	if err != nil {
		return err

	}
	return err
}

func (dr *DocxReader) FillFromXml(minSize int) (int, error) {
	var err error
	var data string
	var dataSize int
	for {
		var t xml.Token
		t, err = dr.decoder.Token()
		if t == nil {
			break
		}
		switch se := t.(type) {

		case xml.StartElement:
			if se.Name.Local == "t" {
				var rowValue string
				rowValue, err = dr.collectValues(se.Name.Local)
				if len(rowValue) > 0 {
					data = strings.Join([]string{data, rowValue}, " ")
				}
			}
			if len(data) >= minSize {
				goto End
			}

		default:

		}
		if err == io.EOF {
			break
		}
	}
End:
	if err == io.EOF {
		dr.err = io.EOF
	}
	byteData := []byte(data)
	dr.data = append(dr.data, byteData...)
	dataSize += len(byteData)
	return dataSize, err
}

func (dr *DocxReader) collectValues(elem string) (rowString string, err error) {
	for {
		var t xml.Token
		t, err = dr.decoder.Token()
		if t == nil {
			break
		}
		switch se := t.(type) {
		case xml.CharData:
			rowString = string([]byte(se))
		case xml.EndElement:
			return rowString, nil
		}
	}
	return rowString, err
}
