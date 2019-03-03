package xlsxparser

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"io"
	"log"
	"strconv"
	"strings"
	"util"
)

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
func Parse(path string) (parsedString string, err error) {
	if !util.FileExists(path) {
		return parsedString, errors.New("file does not exist")
	}
	var r *zip.ReadCloser
	r, err = zip.OpenReader(path)
	defer r.Close()
	if err != nil || r == nil {
		return parsedString, err
	}
	// Iterate through the files in the archive,

	file, sharedStringFile := util.RetrieveWorkBook(r.File)

	if file == nil || sharedStringFile == nil {
		return parsedString, errors.New("file is not valid xlsx")
	}

	rc, err := file.Open()
	if err != nil {
		return "", err

	}
	data, _ := util.ReadFile(rc)

	if err != nil {
		return "", err
	}

	rc, err = sharedStringFile.Open()
	if err != nil {
		return "", err

	}

	strData, _ := util.ReadFile(rc)

	var sr sharedStrings

	err = xml.Unmarshal([]byte(strData), &sr)

	if err != nil {
		log.Println(err)

	}

	byteValue := []byte(data)

	sheets := sheetNumbers{}

	err = xml.Unmarshal(byteValue, &sheets)
	values := []string{}
	for key := range sheets.SheetNumber.Elems {
		values = append(values, key)
		f := util.RetrieveSheetWithNumber(r.File, key)
		if f != nil {
			rcc, err := f.Open()
			if err != nil {
				continue
			}
			sharedStringsData, _ := util.ReadFile(rcc)
			var shd sheetsData
			err = xml.Unmarshal([]byte(sharedStringsData), &shd)

			if err != nil {
				continue

			}
			for _, sheetData := range shd.SheetData {
				for i, row := range sheetData.Row {
					rowString := ""
					for _, col := range row.Col {
						if col.Key == "inlineStr" {
							sa := []string{rowString, col.InlineStr.Value}
							rowString = strings.Join(sa, " ")
						}
						if col.Key == "n" || col.Key == "" {
							sa := []string{rowString, col.Value}
							rowString = strings.Join(sa, " ")
						}
						if col.Key == "s" {
							strngIndex, _ := strconv.ParseInt(col.Value, 10, 64)
							sa := []string{rowString, sr.SharedString[strngIndex].Text}
							rowString = strings.Join(sa, " ")
						}
					}
					delimiter := "\n"
					if i == 0 {
						delimiter = ""
					}
					parsedString = strings.Join([]string{parsedString, rowString}, delimiter)
				}
			}
		}
	}

	return parsedString, err
}
