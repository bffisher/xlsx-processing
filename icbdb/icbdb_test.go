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

func Test_readConfig(t *testing.T) {
	test_data = &data_t{}
	test_data.path = "../test_files/icbdb/"
	conf, err := readConfig(test_data.path + "_conf.xlsx")
	if err != nil {
		t.Fatal(err)
	}

	test_data.conf = conf
}
func Test_OpenXlsx(t *testing.T) {
	log.Print("Reading ... ")
	xlsx, err := excelize.OpenFile(test_data.path + test_data.conf.files["OD"])
	if err != nil {
		t.Fatal(err)
	}
	test_data.odXlsx = xlsx

	test_data.odCopaRows = xlsx.GetRows(test_data.conf.sheets["OD_COPA"])
	if len(test_data.odCopaRows) < 2 {
		t.Fatal(errors.New("Can not find 'COPA original data' sheet"))
	}
}
func Test_handleHeader(t *testing.T) {
	odHeader, err := getODCopaHeader(test_data.odCopaRows[0], test_data.conf)
	if err != nil {
		t.Fatal(err)
	}
	test_data.odCopaHeader = odHeader
	test_data.odCopaRows = test_data.odCopaRows[1:]
	log.Println("OK!")
}

func Test_splitIcbDb(t *testing.T) {
	log.Print("Matching ... ")
	splitIcbDb(test_data)
}

func Test_resolveIcbDbRelation(t *testing.T) {
	err := resolveIcbDbRelation(test_data)
	if err != nil {
		t.Fatal(err)
	}
}

func Test_handleUnusedDb(t *testing.T) {
	handleUnusedDb(test_data)
	log.Println("OK!")
}

func Test_result(t *testing.T) {
	log.Print("Outputing ... ")
	output(test_data)
	log.Println("OK!")
}
