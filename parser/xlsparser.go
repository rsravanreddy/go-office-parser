package parser

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"math"
	"os"
	"strings"

	"github.com/richardlehane/mscfb"
	"github.com/rsravanreddy/go-office-parser/util"
)

type XlsReader struct {
	sheets        []BoundSheet
	sharedStrings []string
	formats       []uint16
	workbook      *mscfb.File
	file          *os.File
	currentSheet  int
	sheetToPos    map[int]uint32
	err           error
	offset        int
	data          []byte
}

func NewXlsReader(path string) *XlsReader {
	xr := &XlsReader{}
	xr.currentSheet = 0
	xr.process(path)
	xr.sheetToPos = make(map[int]uint32, len(xr.sheets))
	for i := 0; i < len(xr.sheets); i++ {
		xr.sheetToPos[i] = xr.sheets[i].FilePos
	}
	return xr
}

func (r *XlsReader) Read(b []byte) (int, error) {

	if (r.err != nil || r.err == io.EOF) && len(r.data) == 0 {
		r.err = io.EOF
		return 0, r.err
	}
	//need to fill
	var err error
	var lenRead int
	if r.offset+len(b) > len(r.data) {
		lenRead, err = r.fill(r.offset + len(b) - len(r.data))
	}
	if err != nil && lenRead == 0 {
		return 0, err
	}
	len := util.Min(len(b), len(r.data)-r.offset)
	copy(b[:], r.data[r.offset:])
	r.data = r.data[len:]
	return len, nil

}

func (r *XlsReader) Close() {
	if r.file != nil {
		r.file.Close()
	}
}

var curRecord = 0

var RECORD_TYPE_BOUND_SHEET uint16 = 133

var RECORD_TYPE_COL uint16 = 513
var RECORD_TYPE_ROW uint16 = 520
var RECORD_TYPE_SST uint16 = 252
var RECORD_TYPE_CONTINUE uint16 = 60
var RECORD_TYPE_EOF uint16 = 10

type BoundSheet struct {
	FilePos   uint32
	SheetType uint8
	Visible   uint8
}

type Colinfo struct {
	First   uint16
	Last    uint16
	width   uint16
	xf      uint16
	flags   uint16
	Notused uint16
}

type Col struct {
	Row uint16
	Col uint16
	Xf  uint16
}

type Row struct {
	Index    uint16
	Fcell    uint16
	Lcell    uint16
	Height   uint16
	Notused  uint16
	Notused2 uint16
	Flags    uint16
	Xf       uint16
}

type ExcelBiffHeader struct {
	id     uint16
	length uint16
}

type ExcelExtendedStringRecord struct {
	offset uint64
	length uint64
}

func (r *XlsReader) process(path string) {
	r.file, _ = os.Open(path)
	doc, err := mscfb.New(r.file)
	if err != nil {
		r.err = err
		return
	}

	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		if entry.Name == "Workbook" {
			r.workbook = entry
			var header1 ExcelBiffHeader
			var savedOffet int64
			for {
				headerBuf := make([]byte, 2)
				_, _ = r.workbook.Read(headerBuf)
				header1.id = binary.LittleEndian.Uint16(headerBuf)
				_, _ = r.workbook.Read(headerBuf)
				header1.length = binary.LittleEndian.Uint16(headerBuf)

				if header1.id == RECORD_TYPE_BOUND_SHEET {
					lsavedOffet, _ := r.workbook.Seek(0, os.SEEK_CUR)
					var boundSheet BoundSheet
					binary.Read(r.workbook, binary.LittleEndian, &boundSheet)
					r.sheets = append(r.sheets, boundSheet)
					lsavedOffet, _ = r.workbook.Seek(lsavedOffet, os.SEEK_SET)
				}

				if header1.id == 0x0e0 {
					lsavedOffet, _ := r.workbook.Seek(0, os.SEEK_CUR)
					r.workbook.Seek(2, os.SEEK_CUR)
					var format uint16
					binary.Read(r.workbook, binary.LittleEndian, &format)
					r.formats = append(r.formats, format)
					lsavedOffet, _ = r.workbook.Seek(lsavedOffet, os.SEEK_SET)
				}

				if header1.id == RECORD_TYPE_ROW {
					lsavedOffet, _ := r.workbook.Seek(0, os.SEEK_CUR)
					var row Row
					binary.Read(r.workbook, binary.LittleEndian, &row)
					lsavedOffet, _ = r.workbook.Seek(lsavedOffet, os.SEEK_SET)
				}

				if header1.id == RECORD_TYPE_SST {
					var record ExcelExtendedStringRecord
					var header2 ExcelBiffHeader
					savedOffet, _ = r.workbook.Seek(0, os.SEEK_CUR)
					record.offset = uint64(savedOffet)
					record.length = uint64(header1.length)
					length := header1.length
					var list []ExcelExtendedStringRecord
					list = append(list, record)

					_, err = r.workbook.Seek(int64(length), os.SEEK_CUR)
					tmpBuf := make([]byte, 2)
					readBytes(r.workbook, 2, os.SEEK_CUR, tmpBuf)
					header2.id = binary.LittleEndian.Uint16(tmpBuf)
					readBytes(r.workbook, 2, os.SEEK_CUR, tmpBuf)
					header2.length = binary.LittleEndian.Uint16(tmpBuf)

					for {
						if header2.id != RECORD_TYPE_CONTINUE {
							break
						}
						record.offset = uint64(savedOffet)
						record.length = uint64(header2.length)
						list = append(list, record)
						tmpBuf := make([]byte, 2)
						readBytes(r.workbook, 2, os.SEEK_CUR, tmpBuf)
						header2.id = binary.LittleEndian.Uint16(tmpBuf)
						readBytes(r.workbook, 2, os.SEEK_CUR, tmpBuf)
						header2.length = binary.LittleEndian.Uint16(tmpBuf)
					}
					r.xlGetExtendedRecordString(r.workbook, list)
					r.workbook.Seek(savedOffet, os.SEEK_SET)

				}
				_, err := r.workbook.Seek(int64(header1.length), os.SEEK_CUR)
				if err != nil {
					break
				}

			}
		}
		if err == io.EOF {
			break
		}
	}

}

func (r *XlsReader) fill(minSize int) (lenRead int, err error) {
	var result string
	for i := r.currentSheet; i < len(r.sheets); i++ {
		r.currentSheet = i
		r.workbook.Seek(int64(r.sheetToPos[i]), os.SEEK_SET)
		var header1 ExcelBiffHeader
		for {
			headerBuf := make([]byte, 2)
			_, err = r.workbook.Read(headerBuf)
			header1.id = binary.LittleEndian.Uint16(headerBuf)
			_, _ = r.workbook.Read(headerBuf)
			header1.length = binary.LittleEndian.Uint16(headerBuf)
			buf := make([]byte, header1.length)
			_, err = r.workbook.Read(buf)
			switch header1.id {
			case 0x0A: //end
				//may be exit the outer loop?
				err = io.EOF
				break
			case 0x208:
				break
			case 0x0BD: //MULRK
				var col Col
				binary.Read(bytes.NewReader(buf), binary.LittleEndian, &col)
				var start uint16
				binary.Read(bytes.NewReader(buf[header1.length-2:]), binary.LittleEndian, &start)
				mulrki := start - col.Col
				for i := uint16(0); i <= mulrki; i++ {
					var format uint16
					binary.Read(bytes.NewReader(buf[4+i*6:]), binary.LittleEndian, &format)
					num := numFromRk(buf[4+i*6+2:])
					result = strings.Join([]string{result, fmt.Sprintf("%0.15g", num)}, " ")
				}
				break

			case 0x0BE: //MULBLANK
				break
			case 0x203: //NUMBER
				var val float64
				binary.Read(bytes.NewReader(buf[6:]), binary.LittleEndian, &val)
				result = strings.Join([]string{result, fmt.Sprintf("%0.15g", val)}, " ")
				break
			case 0x06: //FORMULA
				var val float64
				binary.Read(bytes.NewReader(buf[6:]), binary.LittleEndian, &val)
				var col Col
				binary.Read(bytes.NewReader(buf), binary.LittleEndian, &col)
				result = strings.Join([]string{result, r.formatCellValue(val, col.Xf)}, " ")
				break
			case 0x27e: //RK
				var val float64
				binary.Read(bytes.NewReader(buf[len(buf)-4:]), binary.LittleEndian, &val)
				result = strings.Join([]string{result, fmt.Sprintf("%0.15g", val)}, " ")
				break
			case 0xFD: //LABELSST
				var val uint32
				binary.Read(bytes.NewReader(buf[6:]), binary.LittleEndian, &val)
				if val < uint32(len(r.sharedStrings)) {
					result = strings.Join([]string{result, r.sharedStrings[val]}, " ")
				}
				break
			default:
				break
			}
			saveOffset, _ := r.workbook.Seek(0, os.SEEK_CUR)
			r.sheetToPos[i] = uint32(saveOffset)

			if err == io.EOF || len(result) >= minSize {
				goto End
			}

		}
	}
End:
	if err == io.EOF {
		r.currentSheet++
	}
	if (err == io.EOF || err != nil) && r.currentSheet >= len(r.sheets) {
		r.err = err
	}
	byteData := []byte(result)
	r.data = append(r.data, byteData...)
	return len(byteData), err
}

func (r *XlsReader) xlGetExtendedRecordString(file *mscfb.File, list []ExcelExtendedStringRecord) {
	var curRecord int
	record := list[curRecord]
	_, err := file.Seek(int64(record.offset), os.SEEK_SET)
	if err != nil {
		return
	}
	tmpBuf := make([]byte, 4)
	file.Read(tmpBuf)
	tmpBuf = make([]byte, 4)
	file.Read(tmpBuf)
	cstUnique := binary.LittleEndian.Uint32(tmpBuf)
	var i uint32
	for {
		var cch uint16
		if i >= cstUnique {
			break
		}
		tmpBuf := make([]byte, 2)
		file.Read(tmpBuf)
		cch = binary.LittleEndian.Uint16(tmpBuf)
		isHighByte, pcRun, pcbExtRst := readExcelStringFlag(file, list)
		if isHighByte {
			cch *= 2
		}

		if !r.readExcelString(file, int64(cch), list) {
			break
		}
		if pcRun > 0 {
			file.Seek(int64(4*pcRun), os.SEEK_CUR)
		}
		if pcbExtRst > 0 {
			file.Seek(int64(pcbExtRst), os.SEEK_CUR)

		}

		i++

	}

}

func (r *XlsReader) readExcelString(file *mscfb.File, chunkSize int64, list []ExcelExtendedStringRecord) bool {
	record := list[curRecord]
	if changeExcelRecordIfNeeded(file, list) && curRecord >= len(list) {
		return false
	}
	curPosition, _ := file.Seek(0, os.SEEK_CUR)
	curRecordEnd := int64(record.offset + record.length)
	if curPosition+chunkSize <= curRecordEnd {
		buf := make([]byte, chunkSize)
		_, err := file.Read(buf)
		s := string(buf)
		r.sharedStrings = append(r.sharedStrings, s)

		if err != nil {
			return false
		}
		return true
	} else if curRecordEnd < curPosition {
		return false
	} else {
		var chunkSizeFirstRecord int64
		var chunkSizeSecondRecord int64
		chunkSizeFirstRecord = curRecordEnd - curPosition
		chunkSizeSecondRecord = chunkSize - chunkSizeFirstRecord
		buf := make([]byte, chunkSizeFirstRecord)
		n, _ := file.Read(buf)
		if n > 0 {
			curRecord++
			if curRecord < len(list) {
				record = list[curRecord]
				file.Seek(int64(record.offset), os.SEEK_SET)
				readExcelStringFlag(file, list)
				buf := make([]byte, chunkSizeSecondRecord)
				n, _ := file.Read(buf)
				if n > 0 {
					return true
				}
			}

		}
		return false
	}

}

func readExcelStringFlag(file *mscfb.File, list []ExcelExtendedStringRecord) (isHighByte bool, pcRun int16, pcbExtRst uint16) {
	var bitMask uint8
	binary.Read(file, binary.LittleEndian, &bitMask)
	isHighByte = (bitMask & 0x01) == 0x01
	isExtstring := (bitMask & 0x04) == 0x04
	isRichString := (bitMask & 0x08) == 0x08

	if isRichString {
		binary.Read(file, binary.LittleEndian, &pcRun)
	}
	if isExtstring {
		buf := make([]byte, 4)
		file.Read(buf)
		pcbExtRst = binary.LittleEndian.Uint16(buf[:2])
	}
	return isHighByte, pcRun, pcbExtRst
}

func changeExcelRecordIfNeeded(file *mscfb.File, list []ExcelExtendedStringRecord) bool {
	record := list[curRecord]
	curPosition, _ := file.Seek(0, os.SEEK_CUR)
	if curPosition >= int64(record.offset+record.length) {
		curRecord++
		if curRecord < len(list) {
			record = list[curRecord]
			file.Seek(int64(record.offset), os.SEEK_SET)
		}
		return true
	}
	return false
}

func numFromRk(buf []byte) float64 {
	var val int32
	binary.Read(bytes.NewReader(buf), binary.LittleEndian, &val)
	var num float64
	if val&0x02 > 0 {
		num = float64(val >> 2)
	} else {
		val := uint64(val >> 2)
		val = val << 34
		num = math.Float64frombits(uint64(val))
	}

	if val&0x01 == 1 {
		num = float64(num / 100.0)
	}
	return num
}

func (r *XlsReader) formatCellValue(f float64, format uint16) string {
	var ret string
	ret = fmt.Sprintf("%f", f)
	if isFloat(f) {
		return ret
	}
	if int(format) < len(r.formats) {
		format = r.formats[format]
	}
	switch format {
	case 0:
		ret = fmt.Sprintf("%d", int(f))
		break
	case 1:
		ret = fmt.Sprintf("%d", int(f))
		break
	case 2:
		ret = fmt.Sprintf("%.1f", f)
		break
	case 9:
		ret = fmt.Sprintf("%d", int(f))
		break
	case 10:
		ret = fmt.Sprintf("%.2f", f)
		break
	case 11:
		ret = fmt.Sprintf("%.1e", f)
		break
	case 14:
		ret = fmt.Sprintf("%.0f", f)
		break
	default:
		ret = fmt.Sprintf("%.2f", f)
		break
	}
	return ret
}

func isFloat(a float64) bool {
	b := float64(int64(a))
	c := math.Abs(a - b)
	if c > 0 {
		return true
	}
	return false
}
