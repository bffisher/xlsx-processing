package icbdb

import (
	"errors"
	"fmt"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const TEST_FILE_PATH = "../test_files/icbdb/"

var test_xlsx *excelize.File
var test_data *data_t

func Test_readConfig(t *testing.T) {
	test_data = &data_t{}
	conf, err := readConfig(TEST_FILE_PATH + "_conf.xlsx")
	if err != nil {
		t.Fatal(err)
	}

	test_data.conf = conf
}
func Test_OpenXlsx(t *testing.T) {
	fmt.Print("Reading ... ")
	xlsx, err := excelize.OpenFile(TEST_FILE_PATH + test_data.conf.files["OD"])
	if err != nil {
		t.Fatal(err)
	}
	test_data.odXlsx = xlsx

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
	odHeader, err := getODHeader(test_data.odCopaRows[0], test_data.odIcbOrdRows[0], test_data.conf)
	if err != nil {
		t.Fatal(err)
	}
	test_data.odHeader = odHeader
	test_data.odCopaRows = test_data.odCopaRows[1:]
	test_data.odIcbOrdRows = test_data.odIcbOrdRows[1:]
	fmt.Println("OK!")
}

func Test_splitIcbDb(t *testing.T) {
	fmt.Print("Calculating ... ")
	splitIcbDb(test_data)
}

func Test_resolveIcbDbRelation(t *testing.T) {
	resolveIcbDbRelation(test_data)
}

func Test_handleUnusedDb(t *testing.T) {
	handleUnusedDb(test_data)
	fmt.Println("OK!")
}

func Test_result(t *testing.T) {
	fmt.Print("Outputing ... ")
	output(test_data, TEST_FILE_PATH)
	fmt.Println("OK!")
}
