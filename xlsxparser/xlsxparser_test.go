package xlsxparser

import (
	"path/filepath"
	"strings"
	"testing"
)

var expectedParsedString = `
 Excel 2007 (xlsx) Sample Worksheet
 Created with Microsoft Excel 2007 SP1
 X Y
 0.71852083941501133 0.91682636398560935
 0.48851347410198098 0.67533605773981398
 0.98275635763881475 0.9756846511453845
 0.59948315997276769 0.19516415275790377
 0.19274700075458306 8.9876073934985534E-2
 0.1738905436072411 0.29243269016944495
 0.25690290155942197 0.24113034931149535
 0.79066209083505568 0.42115452316288415
 0.55892707996575974 0.52808780257052512
 3.6696900125545717E-2 0.74098916307003471
 0.4348288110814984 0.5857189400634768
 0.92564680637977315 0.57859457551440485
 8.5392279055289677E-2 0.62102022860769712
 0.42413449198404973 0.5112230103470714
 0.19939336884031489 0.83319968566685776
 0.24813032540125546 0.68208880193637622
 0.3556405079164342 0.1471238009164697`

func TestParse(t *testing.T) {
	path, _ := filepath.Abs("../../testdata/sample.xlsx")
	parsedString, err := Parse(path)
	if err == nil {
		if strings.Compare(expectedParsedString, parsedString) != 0 {
			t.Error(expectedParsedString, parsedString)

		}
	} else {
		t.Error(nil, err)
	}
}

func TestParseInvalidXLSX(t *testing.T) {
	_, err := Parse("/Users/sravanrekula/Downloads/invalid_xlsx.xlsx")
	if err == nil {
		t.Fail()
	}
}

func TestParseNonExistentXLSX(t *testing.T) {
	_, err := Parse("/Users/sravanrekula/Downloads/nofile.xlsx")
	if err == nil {
		t.Fail()
	}
}
