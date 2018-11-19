package unbilledCost
import (
	"errors"
	"log"
	//"strconv"
	"strings"
	"xlsx-processing/util"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const DB_IDX_NO_RELATION = -1
const DB_IDX_NO_DATA = -2
const WBSNO_LEN = 17
const RIGHT_CUSTOMER_PREFIX = "P"

type uc_left_data_t struct{
	idx int
	rightIdxs []int
}

type uc_right_data_t struct{
	idx int
	leftIdx int
}

var data struct {
	conf *config_t
	ucXlsx, ioXlsx, fgXlsx, sgXlsx *excelize.File
	ucHeader, ioHeader map[string]int
	ucData [][]string
	ucLeftData []uc_left_data_t
	ucRightData []uc_right_data_t
}

func Exec(confFile string) error {
	var err error
	data.conf, err = readConfig(confFile)
	if(err != nil) {return err}

	err=openExcelFiles();
	if(err != nil) {return err}
	
	err = readUC()
	if(err != nil) {return err}

	err = splitUC()
	if(err != nil) {return err}

	err = resolveUCLeftRightRelation()
	if(err != nil) {return err}
	
	log.Println("OK!")
	return nil
}

func openExcelFiles() error{
	var err error
	log.Print("Open excel files... ")
	
	data.ucXlsx, err = excelize.OpenFile(data.conf.files["UC_FILE"])
	if err != nil {
		return err
	}
	data.ioXlsx, err = excelize.OpenFile(data.conf.files["IO_FILE"])
	if err != nil {
		return err
	}
	data.fgXlsx, err = excelize.OpenFile(data.conf.files["FG_FILE"])
	if err != nil {
		return err
	}
	data.sgXlsx, err = excelize.OpenFile(data.conf.files["SG_FILE"])
	if err != nil {
		return err
	}
	return nil
}

func readUC()error{
	rows:=data.ucXlsx.GetRows(data.conf.sheets["UC_SHEET"]);

	data.ucHeader = make(map[string]int)
	err := handleHeader(rows[0], data.ucHeader, "UC_")
	if(err != nil) {return err}
	
	data.ucData = rows[1:]

	ucObjColIdx := data.ucHeader["UC_OBJ"]
	for index, row := range data.ucData {
		data.ucData[index][ucObjColIdx] = parseUCObjValue(strings.TrimSpace(row[ucObjColIdx]))
	}
	return nil
}

func splitUC()error{
	len := len(data.ucData)
	data.ucLeftData = make([]uc_left_data_t, 0, len/2)
	data.ucRightData = make([]uc_right_data_t, 0, len/3*2)

	customerIdx := data.ucHeader["UC_CUST"]
	for index, row := range data.ucData {
		if strings.Index(strings.ToUpper(row[customerIdx]), RIGHT_CUSTOMER_PREFIX) == 0 {
			data.ucRightData = append(data.ucRightData, uc_right_data_t{index, DB_IDX_NO_RELATION})
		}else{
			data.ucLeftData = append(data.ucLeftData, uc_left_data_t{index, make([]int, 0)})
		}
	}

	return nil
}

func resolveUCLeftRightRelation()error{
	err := findUCRightByIcbOrd()
	if(err != nil) {return err}
	return nil
}

func findUCRightByIcbOrd()error{
	sheet := data.conf.sheets["IO_SHEET"]
	rows := data.ioXlsx.GetRows(sheet)
	if len(rows) < 2 {
		return errors.New("Can not find '" + sheet + "' sheet")
	}

	data.ioHeader = make(map[string]int)
	err := handleHeader(rows[0], data.ioHeader, "IO_")
	if(err != nil) {return err}

	rows = rows[1:]

	soNoColIdx := data.ioHeader["IO_SO_NO"]
	wbsColIdx := data.ioHeader["IO_WBS"]
	dbSoNoColIdx := data.ioHeader["IO_DB_SO_NO"]
	for index, row := range rows {
		rows[index][wbsColIdx] = parseWbsNoValue(strings.TrimSpace(row[wbsColIdx]))
	}
	err = findUCRight(rows, soNoColIdx, wbsColIdx, dbSoNoColIdx)
	if(err != nil) {return err}

	return nil
}

func handleHeader(row []string, header map[string]int, prefix string)error{
	filter := func(key string) bool {
		return strings.Index(key, prefix) == 0
	}
	util.ConvColNameToIdx(row, data.conf.columns, header, filter)
	err := util.CheckColNameIdx(data.conf.columns, header, filter)
	if(err != nil){
		return err
	}

	return nil
}

func findUCRight(rows [][]string, soNoColIdx, wbsColIdx, dbSoNoColIdx int)error{
	ucObjColIdx := data.ucHeader["UC_OBJ"]
	for leftIndex, left := range data.ucLeftData{
		if len(left.rightIdxs) > 0{
			continue
		}

		ucObjVal := data.ucData[left.idx][ucObjColIdx]

		for _, row := range rows{
			wbsVal := row[wbsColIdx]
			soNoVal := strings.TrimSpace(row[soNoColIdx])
			if(wbsVal != "" && ucObjVal == wbsVal || ucObjVal ==  soNoVal){
				dbSoNoVal := strings.TrimSpace(row[dbSoNoColIdx])
				for rightIndx, right := range data.ucRightData{
					if right.leftIdx != DB_IDX_NO_RELATION && right.leftIdx != DB_IDX_NO_DATA{
						continue
					}

					if dbSoNoVal == data.ucData[right.idx][ucObjColIdx]{
						data.ucLeftData[leftIndex].rightIdxs = append(data.ucLeftData[leftIndex].rightIdxs, right.idx)
						data.ucRightData[rightIndx].leftIdx = left.idx
					}
				}
			}
		}
	}
	return nil
}

func parseUCObjValue(val string) string{
	index := strings.Index(val, "/")
	if index >= 0{
		return val[0: index]
	}

	return parseWbsNoValue(val)
}

func parseWbsNoValue(val string)string{
	if len(val) > WBSNO_LEN{
		return val[0:WBSNO_LEN]
	}

	return val
}