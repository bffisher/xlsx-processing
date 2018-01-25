package copa

import (
	"strings"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const TEST_FILE_PATH = "../test_files/copatest.xlsx"
const TEST_SHEET_NAME = "by p10"

var test_xlsx *excelize.File
var test_header *header_t

func Init() {

}

func Test_OpenXlsx(t *testing.T) {
	var err error
	test_xlsx, err = excelize.OpenFile(TEST_FILE_PATH)
	if err != nil {
		t.Fatal(err)
	}
}

func Test_handleHeader(t *testing.T) {
	header, err := handleHeader(test_xlsx, TEST_SHEET_NAME)
	if err != nil {
		t.Fatal(err)
	}

	if header.hierarchyIdx < 0 {
		t.Fatal("Can not find hierarchy column", header)
	}

	if header.businessIdx < 0 {
		t.Fatal("Can not find business column", header)
	}

	if header.tradPartnIdx < 0 {
		t.Fatal("Can not find trad partn column", header)
	}

	if header.exportIdx < 0 {
		t.Fatal("Can not find export column", header)
	}

	if header.profitCenterIdx < 0 {
		t.Fatal("Can not find profit center column", header)
	}

	if header.partnerProfitCenterIdx < 0 {
		t.Fatal("Can not find partner profit center column", header)
	}

	test_header = &header
}

func Test_handleBody(t *testing.T) {
	handleBody(test_xlsx, TEST_SHEET_NAME, test_header)
}

func Test_End(t *testing.T) {
	err := test_xlsx.SaveAs(strings.Replace(TEST_FILE_PATH, ".xlsx", "_res.xlsx", -1))
	if err != nil {
		t.Fatal(err)
	}
}
