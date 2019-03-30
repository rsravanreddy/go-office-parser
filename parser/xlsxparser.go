package parser

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"io"
	"log"
	"strconv"
	"strings"

	"github.com/rsravanreddy/go-office-parser/util"
)

type XlsxReader struct {
	err                error
	data               []byte
	offset             int
	length             int
	sharedStringsValue sharedStrings
	sheetNumberValue   int
	rowValue           int
	colValue           int
	key                int
	shd                *sheetsData
	zipReader          *zip.ReadCloser
	numberOfSheets     int
	sheetDecoder       *xml.Decoder
	sheetRc            io.ReadCloser
}

func NewXlsxReader(path string) (*XlsxReader, error) {
	dr := &XlsxReader{}
	dr.offset = 0
	dr.length = 0
	dr.key = 1
	dr.err = dr.parse(path)
	return dr, dr.err
}

func (r *XlsxReader) Read(b []byte) (int, error) {

	if r.err != nil || r.err == io.EOF {
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

func (r *XlsxReader) Close() error {
	r.err = errors.New("reader already closed")
	r.sharedStringsValue.SharedString = nil
	if r.zipReader != nil {
		r.zipReader.Close()
	}
	if r.sheetRc != nil {
		r.sheetRc.Close()
	}
	return nil
}

type sheetNumbers struct {
	SheetNumber sheetNumber `xml:"sheets"`
}

type sheetNumber struct {
	Elems map[string]string
}

type sharedStrings struct {
	SharedString []struct {
		Text string `xml:"t"`
	} `xml:"si"`
}

type sheetsData struct {
	SheetData []struct {
		Row []struct {
			Col []struct {
				Key       string `xml:"t,attr"`
				Index     string `xml:"s,attr"`
				Value     string `xml:"v"`
				InlineStr struct {
					Value string `xml:"t"`
				} `xml:"is"`
			} `xml:"c"`
		} `xml:"row"`
	} `xml:"sheetData"`
}

func (sn *sheetNumber) UnmarshalXML(d *xml.Decoder, start xml.StartElement) (err error) {
	type entry struct {
		Key   string `xml:"sheetId,attr"`
		Value string `xml:",chardata"`
	}
	e := entry{}
	sn.Elems = map[string]string{}
	for err = d.Decode(&e); err == nil; err = d.Decode(&e) {
		sn.Elems[e.Key] = e.Value
	}
	if err != nil && err != io.EOF {
		return err
	}
	return nil
}

//Parse .. paerses an excel file and returns as a formatted string
func (dr *XlsxReader) parse(path string) (err error) {
	if !util.FileExists(path) {
		return errors.New("file does not exist")
	}
	dr.zipReader, err = zip.OpenReader(path)
	if err != nil || dr.zipReader == nil {
		return err
	}
	// Iterate through the files in the archive,

	file, sharedStringFile := util.RetrieveWorkBook(dr.zipReader.File)

	if file == nil || sharedStringFile == nil {
		return errors.New("file is not valid xlsx")
	}

	rc, err := file.Open()
	if rc != nil {
		defer rc.Close()
	}
	if err != nil {
		return err

	}

	data, _ := util.ReadFile(rc)

	if err != nil {
		return err
	}

	src, err := sharedStringFile.Open()
	if src != nil {
		defer src.Close()
	}
	if err != nil {
		return err

	}

	strData, _ := util.ReadFile(src)

	err = xml.Unmarshal([]byte(strData), &dr.sharedStringsValue)

	if err != nil {
		log.Println(err)

	}

	byteValue := []byte(data)

	sheets := sheetNumbers{}

	err = xml.Unmarshal(byteValue, &sheets)
	dr.numberOfSheets = len(sheets.SheetNumber.Elems)

	return err
}

func (dr *XlsxReader) FillFromXml(minSize int) (dataSize int, err error) {
	var totalSizeRead int
	for key := dr.key; key <= dr.numberOfSheets; key++ {
		dr.key = key
		sizeRead, err := dr.parseSheet(minSize - totalSizeRead)
		totalSizeRead += sizeRead
		if totalSizeRead >= minSize {
			break
		}
		if err != nil {
			continue
		}
		dr.sheetNumberValue = key
	}
	return totalSizeRead, err
}

func (dr *XlsxReader) parseSheet(minSize int) (dataSize int, err error) {
	var rowString string
	if dr.sheetDecoder == nil {
		f := util.RetrieveSheetWithNumber(dr.zipReader.File, strconv.Itoa(dr.key))
		dr.sheetRc, err = f.Open()
		if err == nil {
			dr.sheetDecoder = xml.NewDecoder(dr.sheetRc)
		}
	}
	if dr.sheetDecoder != nil {
		//var inElement string

		for {
			var t xml.Token
			t, err = dr.sheetDecoder.Token()
			if t == nil {
				break
			}
			switch se := t.(type) {
			case xml.StartElement:
				if se.Name.Local == "row" {
					var rowValue string
					rowValue, err = dr.collectCols(se.Name.Local)
					rowString = strings.Join([]string{rowString, rowValue}, "")

				}
				if len(rowString) >= minSize {
					goto End
				}
				//}
			default:
			}
			if err == io.EOF {
				dr.sheetDecoder = nil
				dr.sheetRc.Close()
				break
			}
		}
	}

End:
	if err == io.EOF {
		dr.sheetDecoder = nil
		dr.sheetRc.Close()
	}
	byteData := []byte(rowString)
	if dr.key == dr.numberOfSheets && err == io.EOF {
		dr.err = io.EOF
	}
	dr.data = append(dr.data, byteData...)
	return dataSize, err
}

func (dr *XlsxReader) collectCols(elem string) (rowString string, err error) {
	for {
		var t xml.Token
		t, err = dr.sheetDecoder.Token()
		if t == nil {
			break
		}
		switch se := t.(type) {
		case xml.StartElement:
			if se.Name.Local == "c" {
				var colType string
				for i := 0; i < len(se.Attr); i++ {
					if se.Attr[i].Name.Local == "t" {
						colType = se.Attr[i].Value
					}
				}
				var colValue string
				colValue, err = dr.collectValues(se.Name.Local, colType)
				rowString = strings.Join([]string{rowString, colValue}, "")
			}
		case xml.EndElement:
			if se.Name.Local == elem {
				return rowString, err
			}

		}
	}

	return rowString, err
}

func (dr *XlsxReader) collectValues(elem string, colType string) (rowString string, err error) {
	for {
		var t xml.Token
		t, err = dr.sheetDecoder.Token()
		if t == nil {
			break
		}
		switch se := t.(type) {
		case xml.StartElement:
			if se.Name.Local == "v" {
				var tt xml.Token
				tt, err = dr.sheetDecoder.Token()
				if tt == nil {
					break
				}
				switch cd := tt.(type) {
				case xml.CharData:
					str := string([]byte(cd))
					//fmt.Println(str)
					if colType == "inlineStr" {
						sa := []string{rowString, str}
						rowString = strings.Join(sa, " ")
					}
					if colType == "n" || colType == "" {
						sa := []string{rowString, str}
						rowString = strings.Join(sa, " ")
					}
					if colType == "s" {
						strngIndex, _ := strconv.ParseInt(str, 10, 64)
						sa := []string{rowString, dr.sharedStringsValue.SharedString[strngIndex].Text}
						rowString = strings.Join(sa, " ")
					}
					//fmt.Println(rowString)
				}
			}
		case xml.EndElement:
			if se.Name.Local == "c" {
				return rowString, err
			}

		}
	}
	return rowString, err
}
