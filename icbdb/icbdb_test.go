package icbdb

import (
	"errors"
	"log"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const TEST_FILE_PATH = "../test_files/icbdb/"

var test_xlsx *excelize.File
var test_data *data_t

func Test_handleConfig(t *testing.T) {
	test_data = &data_t{}
	conf, err := handleConfig(TEST_FILE_PATH + "_conf.xlsx")
	if err != nil {
		t.Fatal(err)
	}

	test_data.conf = &conf
}
func Test_OpenXlsx(t *testing.T) {
	xlsx, err := excelize.OpenFile(TEST_FILE_PATH + test_data.conf.files["OD_FILE"])
	if err != nil {
		t.Fatal(err)
	}
	test_data.odCopaRows = xlsx.GetRows(test_data.conf.sheets["OD_COPA"])
	if len(test_data.odCopaRows) < 2 {
		t.Fatal(errors.New("Can not find 'COPA original data' sheet"))
	}
	test_data.odIcbOrdRows = xlsx.GetRows(test_data.conf.sheets["OD_ICB_ORD"])
	if len(test_data.odIcbOrdRows) < 2 {
		t.Fatal(errors.New("Can not find 'ICB_ORD' sheet"))
	}
}
func Test_handleHeader(t *testing.T) {
	odCopaHeader, err := handleODCopaHeader(test_data.odCopaRows[0], test_data.conf)
	if err != nil {
		t.Fatal(err)
	}
	test_data.odCopaHeader = &odCopaHeader
	test_data.odCopaRows = test_data.odCopaRows[1:]

	odIcbOrdHeader, err := handleODIcbOrdHeader(test_data.odIcbOrdRows[0], test_data.conf)
	if err != nil {
		t.Fatal(err)
	}
	test_data.odIcbOrdHeader = &odIcbOrdHeader
	test_data.odIcbOrdRows = test_data.odIcbOrdRows[1:]
}

func Test_splitIcbDb(t *testing.T) {
	splitIcbDb(test_data)
}

func Test_resolveIcbDbRelation(t *testing.T) {
	resolveIcbDbRelation(test_data)
}

func Test_result(t *testing.T) {
	for key, value := range test_data.icbDbRelation {
		if value < 0 {
			no := test_data.odCopaRows[key][test_data.odCopaHeader.soNoIdx]
			log.Println(no, value)
		}
	}

}
