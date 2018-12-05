package wip

import (
	"errors"
	"log"
	//"strings"
	"xlsx-processing/util"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type config_t struct {
	files, sheets map[string]string
	gis2Sheets            [][2]string
}

var data struct {
	conf *config_t
	wipXlsx, gisXlsx, gis2Xlsx *excelize.File
}

func Exec()error{
	var err error
	data.conf, err = readConfig(util.Env().ConfigFullName)
	if(err != nil) {return err}

	err = openExcelFiles(util.Env().FilePath);
	if(err != nil) {return err}

	log.Println("Read wip data...")
	rows := data.wipXlsx.GetRows(data.conf.sheets["WIP_SHEET"]);

	wipHeader := rows[0]	
	//wipData := rows[1:]

	wipOrderTypeColIdx := -1
	salesOrderColIdx := -1
	wbsColIdx := -1
	projectNameColIdx := -1
	customerNameColIdx := -1
	for index, name := range wipHeader {
		if name == "Order Type" {
			wipOrderTypeColIdx = index
		}else if name == "Sales Order"{
			salesOrderColIdx = index
		}else if name == "WBS Element"{
			wbsColIdx = index
		}else if name == "Project Name"{
			projectNameColIdx = index
		}else if name == "Customer Name"{
			customerNameColIdx = index
		}
	}

	if wipOrderTypeColIdx < 0 {
		return errors.New("Can't find Order Type column!")
	}else if salesOrderColIdx < 0{
		return errors.New("Can't find Sales Order column!")
	}else if wbsColIdx < 0{
		return errors.New("Can't find WBS Elemente column!")
	}else if projectNameColIdx < 0{
		return errors.New("Can't find Project Name column!")
	}else if customerNameColIdx < 0{
		return errors.New("Can't find 客户名称 column!")
	}



	return nil
}

func readConfig(filePath string) (*config_t, error) {
	conf := config_t{make(map[string]string), make(map[string]string), make([][2]string, 0, 10)}

	rows, err := util.ReadConfig(filePath)
	if err != nil {
		return &conf, err
	}

	for _, row := range rows {
		util.CopyRowToMap(conf.files, row, 0, 1)
		util.CopyRowToMap(conf.sheets, row, 3, 4)
		conf.gis2Sheets = util.CopyRowToArray(conf.gis2Sheets, row, 6, 7)
	}

	return &conf, nil
}

func openExcelFiles(filePath string) error{
	var err error
	log.Print("Open excel files... ")
	
	data.wipXlsx, err = excelize.OpenFile(filePath + data.conf.files["WIP_FILE"])
	if err != nil {
		return err
	}
	data.gisXlsx, err = excelize.OpenFile(filePath + data.conf.files["GIS_FILE"])
	if err != nil {
		return err
	}
	data.gis2Xlsx, err = excelize.OpenFile(filePath + data.conf.files["GIS2_FILE"])
	if err != nil {
		return err
	}

	log.Println("Open excel files OK!")
	return nil
}