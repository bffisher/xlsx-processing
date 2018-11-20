package unbilledCost
import (
	"os"
	"errors"
	"log"
	"strconv"
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
	ucXlsx, ioXlsx, gisXlsx, gis2Xlsx *excelize.File
	ucHeader, ioHeader, gisHeader map[string]int
	ucData [][]string
	ucLeftData []uc_left_data_t
	ucRightData []uc_right_data_t
}

func Exec(confFile string) error {
	filePath := os.Getenv("_SOURCE_FILE_PATH")
	var err error
	data.conf, err = readConfig(filePath + confFile)
	if(err != nil) {return err}

	err=openExcelFiles(filePath);
	if(err != nil) {return err}
	
	err = readUC()
	if(err != nil) {return err}

	err = splitUC()
	if(err != nil) {return err}

	err = resolveUCLeftRightRelation()
	if(err != nil) {return err}

	return nil
}

func openExcelFiles(filePath string) error{
	var err error
	log.Print("Open excel files... ")
	
	data.ucXlsx, err = excelize.OpenFile(filePath + data.conf.files["UC_FILE"])
	if err != nil {
		return err
	}
	data.ioXlsx, err = excelize.OpenFile(filePath + data.conf.files["IO_FILE"])
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

func readUC()error{
	log.Println("Read unbilled cost data...")
	rows:=data.ucXlsx.GetRows(data.conf.sheets["UC_SHEET"]);

	data.ucHeader = make(map[string]int)
	err := handleHeader(rows[0], data.ucHeader, "UC_")
	if(err != nil) {return err}
	
	data.ucData = rows[1:]

	ucObjColIdx := data.ucHeader["UC_OBJ"]
	for index, row := range data.ucData {
		data.ucData[index][ucObjColIdx] = parseUCObjValue(strings.TrimSpace(row[ucObjColIdx]))
	}
	log.Println("Read unbilled cost data. OK!")
	return nil
}

func splitUC()error{
	log.Println("Split unbilled cost data...")
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
	log.Println("Split unbilled cost data.OK!")
	return nil
}

func resolveUCLeftRightRelation()error{
	log.Print("Matching... ")
	err := findUCRightByIcbOrd()
	if(err != nil) {return err}

	err = findUCRightByGIS()
	if(err != nil) {return err}

	err = findUCRightByGIS2()
	if(err != nil) {return err}

	log.Println("Matching OK!")
	return nil
}

func findUCRightByIcbOrd()error{
	log.Println("ICB_ORD...")
	sheet := data.conf.sheets["IO_SHEET"]
	rows := data.ioXlsx.GetRows(sheet)
	if len(rows) <= 1 {
		return errors.New("Can not find '" + sheet + "' sheet")
	}

	data.ioHeader = make(map[string]int)
	err := handleHeader(rows[0], data.ioHeader, "IO_")
	if(err != nil) {return err}

	soNoColIdx := data.ioHeader["IO_SO_NO"]
	wbsColIdx := data.ioHeader["IO_WBS"]
	dbSoNoColIdx := data.ioHeader["IO_DB_SO_NO"]

	rows = rows[1:]
	for index, row := range rows {
		rows[index][soNoColIdx] = strings.TrimSpace(row[soNoColIdx])
		rows[index][wbsColIdx] = parseWbsNoValue(strings.TrimSpace(row[wbsColIdx]))
		rows[index][dbSoNoColIdx] = strings.TrimSpace(row[dbSoNoColIdx])
	}
	err = findUCRight(rows, soNoColIdx, wbsColIdx, dbSoNoColIdx)
	if(err != nil) {return err}
	log.Println("ICB_ORD..OK!")
	return nil
}

func findUCRightByGIS()error{
	log.Println("GIS...")
	sheet := data.conf.sheets["GIS_SHEET"]
	headerLineNo := 5
	rows := data.gisXlsx.GetRows(sheet)
	if len(rows) <= headerLineNo {
		return errors.New("Can not find '" + sheet + "' sheet header")
	}

	data.gisHeader = make(map[string]int)
	err := handleHeader(rows[headerLineNo - 1], data.gisHeader, "GIS_")
	if(err != nil) {return err}

	soNoColIdx := data.gisHeader["GIS_SO_NO"]
	wbsColIdx := data.gisHeader["GIS_WBS"]
	dbSoNoColIdx := data.gisHeader["GIS_DB_SO_NO"]

	rows = rows[headerLineNo:]
	for index, row := range rows {
		rows[index][soNoColIdx] = strings.TrimSpace(row[soNoColIdx])
		rows[index][wbsColIdx] = parseWbsNoValue(strings.TrimSpace(row[wbsColIdx]))
		rows[index][dbSoNoColIdx] = strings.TrimSpace(row[dbSoNoColIdx])
	}
	err = findUCRight(rows, soNoColIdx, wbsColIdx, dbSoNoColIdx)
	if(err != nil) {return err}
	log.Println("GIS..OK!")
	return nil
}

func findUCRightByGIS2()error{
	log.Println("GIS2...")
	for _, item := range data.conf.gis2Sheets{
		sheet, headerLineNoStr := item[0], item[1]
		log.Println(sheet + "...")
		headerLineNo, err := strconv.Atoi(headerLineNoStr)
		if(err != nil) {
			log.Printf("Header Line No error(%s)\n", headerLineNoStr)
			continue
		}

		rows := data.gis2Xlsx.GetRows(sheet)
		if len(rows) <= headerLineNo {
			log.Printf("Can not find '%s' sheet\n", sheet)
			continue
		}

		soNoColIdx, wbsColIdx, dbSoNoColIdx:= -1, -1, -1
		for index, name := range rows[headerLineNo - 1] {
			if name == "" {
				continue
			}
			name = strings.ToLower(name)

			if soNoColIdx == -1 && strings.Contains(name, "operation") {
				soNoColIdx = index
			} else if wbsColIdx == -1 && strings.Contains(name, "ccm--wbs") {
				wbsColIdx = index
			} else if dbSoNoColIdx == -1 && strings.Contains(name, "segment") {
				dbSoNoColIdx = index
			}
		}

		if soNoColIdx == -1 {
			return errors.New("Can not find Operation No column in '" + sheet + "' sheet")
		}
		if dbSoNoColIdx == -1 {
			return errors.New("Can not find Segment No column in '" + sheet + "' sheet")
		}

		rows = rows[headerLineNo:]
		for index, row := range rows {
			rows[index][soNoColIdx] = strings.TrimSpace(row[soNoColIdx])
			rows[index][dbSoNoColIdx] = strings.TrimSpace(row[dbSoNoColIdx])
			if wbsColIdx >= 0 {
				rows[index][wbsColIdx] = parseWbsNoValue(strings.TrimSpace(row[wbsColIdx]))
			}
		}
		err = findUCRight(rows, soNoColIdx, wbsColIdx, dbSoNoColIdx)
		if(err != nil) {
			log.Printf("Find in '%s' error\n", sheet)
			continue
		}
		log.Println(sheet + "..OK!")
	}
	log.Println("GIS2..OK!")
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
		emptyCount := 0
		for _, row := range rows{
			wbsVal := ""
			if(wbsColIdx >= 0){
				wbsVal = row[wbsColIdx]
			}
			soNoVal := strings.TrimSpace(row[soNoColIdx])
			if soNoVal == ""{
				//连续100个空值，则认为已到末尾
				emptyCount++
				if emptyCount > 100 {
					break;
				}
			}else{
				emptyCount = 0
			}
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