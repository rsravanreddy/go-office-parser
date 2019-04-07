package parser

import (
	"encoding/binary"
	"errors"
	"io"
	"os"

	"github.com/richardlehane/mscfb"
	"github.com/rsravanreddy/go-office-parser/util"
)

type DocReader struct {
	err         error
	fileMap     map[string]*mscfb.File
	workbook    *mscfb.File
	pieceOffset uint32
	file        *os.File
	pieceCount  uint32
	offset      int
	PieceTable  []byte
	data        []byte
}

func NewDocReader(path string) *DocReader {
	dr := &DocReader{}
	dr.pieceOffset = 0
	dr.fileMap = make(map[string]*mscfb.File)
	dr.process(path)
	return dr
}

func (r *DocReader) Read(b []byte) (int, error) {

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
	//r.offset = r.offset + len
	r.data = r.data[len:]
	return len, nil

}

func (r *DocReader) Close() {
	if r.file != nil {
		r.file.Close()
	}
}

func (dr *DocReader) process(path string) {
	dr.file, _ = os.Open(path)
	//defer file.Close()
	doc, err := mscfb.New(dr.file)
	if err != nil {
		dr.err = err
		return
	}
	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		dr.fileMap[entry.Name] = entry

		if entry.Name == "WordDocument" {
			dr.workbook = entry
			for {
				entry.Seek(0, 0)
				var i uint16
				err := binary.Read(entry, binary.LittleEndian, &i)

				if i != 0xA5EC {
					dr.err = errors.New("invalid doc file")
					return
				}

				//is encrypted
				encbuf := make([]byte, 1)
				entry.Seek(11, 0)
				_, err = entry.Read(encbuf)
				if (encbuf[0] & 0x1) == 0x1 {
					dr.err = errors.New("encrypted file")
					return
				}

				//table stream
				i, err = read16Bit(entry, 0x000A, 0)

				var tableName string
				if (i & 0x0200) == 0x0200 {
					tableName = "1Table"
				} else {
					tableName = "0Table"
				}

				/****** piece table *********/
				ptb := make([]byte, 4)
				_, err = readBytes(entry, 418, 0, ptb)
				fcClx := binary.LittleEndian.Uint32(ptb)
				_, err = readBytes(entry, 418, 1, ptb)
				lcbClx := binary.LittleEndian.Uint32(ptb)
				if lcbClx <= 0 {
					dr.err = errors.New("invalid doc file")
					return
				}
				/****** piece table *********/

				/****** table strem *********/
				tableStreamBuf := make([]byte, lcbClx)
				_, err = readBytes(dr.fileMap[tableName], int64(fcClx), 0, tableStreamBuf)
				i = 0
				var lcbPieceTable uint32
				for {
					if tableStreamBuf[i] == 0x02 {
						lcbPieceTable = binary.LittleEndian.Uint32(tableStreamBuf[i+1:])
						dr.PieceTable = tableStreamBuf[i+5:]
						dr.pieceCount = (lcbPieceTable - 4) / 12
						break
					} else if tableStreamBuf[i] == 1 {
						grpPrlLen := binary.LittleEndian.Uint16(tableStreamBuf[i+1:])
						i = i + 3 + grpPrlLen
					} else {
						break
					}
				}
				if err == io.EOF || err == nil {
					break
				}

			}

		}

	}
}

func (dr *DocReader) fill(minSize int) (lenRead int, err error) {
	//read the word document
	var text []byte
	for {
		var pieceDescriptor []byte
		if dr.pieceOffset >= dr.pieceCount {
			break
		}
		pieceStart := binary.LittleEndian.Uint32(dr.PieceTable[dr.pieceOffset*4:])
		pieceEnd := binary.LittleEndian.Uint32(dr.PieceTable[(dr.pieceOffset+1)*4:])
		pieceDescriptor = dr.PieceTable[((dr.pieceCount+1)*4)+(dr.pieceOffset*8):]
		fc := binary.LittleEndian.Uint32(pieceDescriptor[2:])
		isAnsi := (fc & 0x40000000) == 0x40000000
		if !isAnsi {
			fc = (fc & 0xBFFFFFFF)
		} else {
			fc = (fc & 0xBFFFFFFF) >> 1
		}
		pieceSize := pieceEnd - pieceStart
		if !isAnsi {
			pieceSize *= 2
		}

		if pieceSize >= 1 {
			textbuffer := make([]byte, pieceSize)
			_, err = readBytes(dr.workbook, int64(fc), 0, textbuffer)
			textbuffer = append([]byte(" "), textbuffer...)
			text = append(text, textbuffer...)
		}
		dr.pieceOffset++
		if len(text) > minSize {
			goto End
		}

	}
End:
	if err == io.EOF {
		dr.err = err
	}
	if dr.pieceOffset >= dr.pieceCount {
		dr.err = io.EOF

	}

	dr.data = append(dr.data, text...)
	return len(text), err
}

func read16Bit(file *mscfb.File, offset int64, seek int) (i uint16, err error) {
	file.Seek(offset, seek)
	err = binary.Read(file, binary.LittleEndian, &i)
	return i, err
}

func readBytes(file *mscfb.File, offset int64, seek int, b []byte) (i int, err error) {
	file.Seek(offset, seek)
	i, err = file.Read(b)
	return i, err
}
