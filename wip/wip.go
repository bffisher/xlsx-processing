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

type resutl_t struct{
	idx int
	product, projectName, customerNo, customerName, contractNo string
}

var data struct {
	conf *config_t
	wipXlsx, gisXlsx, gis2Xlsx *excelize.File
	wipData [][]string
	wipHeaderRowIdx int
	wipColIdx util.Col_index_t
	wipOrderTypeColIdx int
	result []resutl_t
}

func Exec()error{
	var err error
	data.conf, err = readConfig(util.Env().ConfigFullName)
	if(err != nil) {return err}

	err = openExcelFiles(util.Env().FilePath);
	if(err != nil) {return err}

	data.result = make([]resutl_t, 0)
	readWip()

	_,gisColIndex, gisRows := util.ReadGIS(data.gisXlsx, data.conf.sheets["GIS_SHEET"])
	find(gisRows, gisColIndex)

	log.Println("GIS2...")
	for _, item := range data.conf.gis2Sheets{
		err,colIdex,rows := util.ReadGIS2Sheet(data.gis2Xlsx, item[0], item[1])
		if(err != nil){
			continue
		}

		find(rows, colIdex)
		log.Println(item[0] + "..OK!")
	}
	log.Println("GIS2..OK!")

	output()

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

func readWip()error{
	log.Println("Read wip data...")
	rows := data.wipXlsx.GetRows(data.conf.sheets["WIP_SHEET"]);

	wipHeader := rows[0]
	data.wipHeaderRowIdx = 0
	data.wipData = rows[1:]
	data.wipColIdx = util.NewColIndex()
	data.wipOrderTypeColIdx = -1
	for index, name := range wipHeader {
		if name == "Order Type" {
			data.wipOrderTypeColIdx = index
		}else if name == "Sales Order"{
			data.wipColIdx.SoNo = index
		}else if name == "WBS Element"{
			data.wipColIdx.Wbs = index
		}else if name == "Product"{
			data.wipColIdx.Product = index
		}else if name == "Project Name"{
			data.wipColIdx.ProjectName = index
		}else if name == "Customer No"{
			data.wipColIdx.CustomerNo = index
		}else if name == "Customer Name"{
			data.wipColIdx.CustomerName = index
		}else if name == "Contract No"{
			data.wipColIdx.ContractNo = index
		}
	}

	if data.wipOrderTypeColIdx < 0 {
		return errors.New("Can't find Order Type column!")
	}else if data.wipOrderTypeColIdx < 0{
		return errors.New("Can't find Sales Order column!")
	}else if data.wipColIdx.Wbs < 0{
		return errors.New("Can't find WBS Elemente column!")
	}
	log.Println("Read wip data...OK!")
	return nil
}

func find(rows [][]string, colIndex util.Col_index_t){
	for wipIdx, wipRow := range(data.wipData){
		if wipRow[data.wipOrderTypeColIdx] != "ZPP5"{
			continue
		}
		wipWbs := util.ParseWbsNoValue(wipRow[data.wipColIdx.Wbs])
		wipSoNo := wipRow[data.wipColIdx.SoNo]
		for _, row := range(rows){
			if wipWbs != "" && colIndex.Wbs > -1 && wipWbs == row[colIndex.Wbs] || wipSoNo == row[colIndex.SoNo]{
				resItem := resutl_t{}
				if colIndex.Product > -1{
					resItem.product = row[colIndex.Product]
				}
				if colIndex.ProjectName > -1{
					resItem.projectName = row[colIndex.ProjectName]
				}
				if colIndex.CustomerNo>-1{
					resItem.customerNo = row[colIndex.CustomerNo]
				}
				if colIndex.CustomerName>-1{
					resItem.customerName = row[colIndex.CustomerName]
				}
				if colIndex.ContractNo>-1{
					resItem.contractNo = row[colIndex.ContractNo]
				}
				resItem.idx = wipIdx + data.wipHeaderRowIdx + 1
				data.result = append(data.result, resItem)
			}
		}
	}
}

func output(){
	log.Println("Output...")

	sheet := data.conf.sheets["WIP_SHEET"]
	for _, resItem := range(data.result){
		if data.wipColIdx.Product > -1 {
			data.wipXlsx.SetCellValue(sheet, util.Axis(resItem.idx, data.wipColIdx.Product), resItem.product)
		}
		if data.wipColIdx.ProjectName > -1 {
			data.wipXlsx.SetCellValue(sheet, util.Axis(resItem.idx, data.wipColIdx.ProjectName), resItem.projectName)
		}
		if data.wipColIdx.CustomerNo > -1 {
			data.wipXlsx.SetCellValue(sheet, util.Axis(resItem.idx, data.wipColIdx.CustomerNo), resItem.customerNo)
		}
		if data.wipColIdx.CustomerName > -1 {
			data.wipXlsx.SetCellValue(sheet, util.Axis(resItem.idx, data.wipColIdx.CustomerName), resItem.customerName)
		}
		if data.wipColIdx.ContractNo > -1 {
			data.wipXlsx.SetCellValue(sheet, util.Axis(resItem.idx, data.wipColIdx.ContractNo), resItem.contractNo)
		}
	}
	data.wipXlsx.Save()
	log.Println("Output...OK!")
}
