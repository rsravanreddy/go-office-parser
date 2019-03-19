package util

import (
	"archive/zip"
	"io"
	"os"
	"strings"
)

func Min(x, y int) int {
	if x < y {
		return x
	}
	return y
}

//ReadFile ...
func ReadFile(r io.Reader) (res string, err error) {
	var sb strings.Builder
	if _, err = io.Copy(&sb, r); err == nil {
		res = sb.String()
	}
	return
}

//RetrieveWordDoc ... retrieve document xml from the archive
func RetrieveWordDoc(files []*zip.File) (file *zip.File) {
	/*
		Simply loops over the files looking for the file with name "word/document"
	*/
	for _, f := range files {
		if f.Name == "word/document.xml" {
			file = f
		}
	}
	return
}

//RetrieveWorkBook ... retrieve sharedstrings xml from the archive

func RetrieveWorkBook(files []*zip.File) (workbook *zip.File, sharedString *zip.File) {
	/*ˆ
	Simply loops over the files looking for the file with name "woxlrd/workbook"
	*/
	for _, f := range files {
		if f.Name == "xl/workbook.xml" {
			workbook = f
		} else if f.Name == "xl/sharedStrings.xml" {
			sharedString = f
		}
	}
	return
}

//RetrieveSheetWithNumber ... retrieve a sheet xml based ont the number of sheet
func RetrieveSheetWithNumber(files []*zip.File, n string) (sheet *zip.File) {
	/*ˆ
	Simply loops over the files looking for the file with name "xl/sheet"
	*/
	for _, f := range files {
		if f.Name == "xl/worksheets/sheet"+n+".xml" {
			sheet = f
		}
	}
	return
}

// FileExists ...
func FileExists(path string) (exists bool) {
	exists = true
	if _, err := os.Stat(path); err != nil {
		if os.IsNotExist(err) {
			exists = false
		}
	}
	return exists
}
