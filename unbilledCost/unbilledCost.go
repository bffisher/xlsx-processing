package unbilledCost
import (
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
	noRightSoNoList []string
	product, projectName, customerNo, customerName, contractNo string
}

type uc_right_data_t struct{
	idx int
	leftIdx int
	dbSoNo string
	product, projectName, customerNo, customerName, contractNo string
}

type col_index_t struct{
	soNo, wbs, dbSoNo int
	product, projectName, customerNo, customerName, contractNo int
}

var data struct {
	conf *config_t
	ucXlsx, ioXlsx, gisXlsx, gis2Xlsx *excelize.File
	ucHeader ,ioHeader, gisHeader []string
	ucData [][]string
	ucObjColIdx, ucCustColIdx int
	ucLeftData []uc_left_data_t
	ucRightData []uc_right_data_t
}

func newColIndex() col_index_t{
	return col_index_t{
		soNo:-1, wbs:-1, dbSoNo:-1,
		product:-1, projectName:-1,customerNo:-1,customerName:-1,contractNo:-1,
	}
}

func Exec() error {
	var err error
	data.conf, err = readConfig(util.Env().ConfigFullName)
	if(err != nil) {return err}

	err=openExcelFiles(util.Env().FilePath);
	if(err != nil) {return err}
	
	err = readUC()
	if(err != nil) {return err}

	err = splitUC()
	if(err != nil) {return err}

	err = resolveUCLeftRightRelation()
	if(err != nil) {return err}

	output()

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

	data.ucHeader = rows[0]	
	data.ucData = rows[1:]

	data.ucObjColIdx = -1
	for index, name := range data.ucHeader {
		if name == "Object" {
			data.ucObjColIdx = index
		}else if name == "Customer"{
			data.ucCustColIdx = index
		}
	}

	if data.ucObjColIdx < 0 {
		return errors.New("Can't find object column!")
	}

	newRows := make([][]string, 0)
	for _, row := range data.ucData {
		orderVal := row[data.ucObjColIdx]
		if strings.Index(orderVal, "800") == 0 || strings.Index(orderVal, "7") == 0{
			continue
		}
		
		row[data.ucObjColIdx] = parseUCObjValue(strings.TrimSpace(row[data.ucObjColIdx]))

		newRows = append(newRows, row)
	}
	data.ucData = newRows

	count := len(data.ucData)
	for i := 0; i < count; i++{
		if data.ucData[i][data.ucObjColIdx] == ""{
			continue
		}

		for j:=i+1; j < count; j++{
			if data.ucData[i][data.ucObjColIdx] ==  data.ucData[j][data.ucObjColIdx]{
				
				for colIdx := data.ucObjColIdx + 3; colIdx < len(data.ucData[j]); colIdx++{
					intval1,_ := strconv.ParseFloat(data.ucData[i][colIdx], 64)
					intval2,_ := strconv.ParseFloat(data.ucData[j][colIdx], 64)
					data.ucData[i][colIdx] = strconv.FormatFloat(intval1 + intval2, 'f', -1, 64)
				}

				data.ucData[j][data.ucObjColIdx] = ""
			}
		}
	}

	newRows = make([][]string, 0)
	for _, row := range data.ucData {
		if row[data.ucObjColIdx] == ""{
			continue
		}
		newRows = append(newRows, row)
	}
	data.ucData = newRows

	log.Println("Read unbilled cost data. OK!")
	return nil
}

func splitUC()error{
	log.Println("Split unbilled cost data...")
	len := len(data.ucData)
	data.ucLeftData = make([]uc_left_data_t, 0, len/2)
	data.ucRightData = make([]uc_right_data_t, 0, len/3*2)

	for index, row := range data.ucData {
		if strings.Index(strings.ToUpper(row[data.ucCustColIdx]), RIGHT_CUSTOMER_PREFIX) == 0 {
			data.ucRightData = append(data.ucRightData, uc_right_data_t{index, DB_IDX_NO_RELATION, "","","","","",""})
		}else{
			data.ucLeftData = append(data.ucLeftData, uc_left_data_t{index, make([]int, 0), make([]string, 0), "","","","",""})
		}
	}
	log.Println("Split unbilled cost data.OK!")
	return nil
}

func resolveUCLeftRightRelation()error{
	log.Print("Matching... ")
	log.Print("Left... ")
	log.Println("ICB_ORD...")
	_ ,icbOrdColIndex, icbOrdRows := readIcbOrd()
	findUCRight(icbOrdRows, icbOrdColIndex)
	log.Println("ICB_ORD..OK!")

	log.Println("GIS...")
	_,gisColIndex, gisRows := readGIS()
	findUCRight(gisRows, gisColIndex)
	log.Println("GIS..OK!")

	findUCRightByGIS2()
	log.Println("Left... OK!")

	log.Print("Right... ")
	log.Println("ICB_ORD...")
	findUCLeft(icbOrdRows, icbOrdColIndex)
	log.Println("ICB_ORD..OK!")
	log.Println("GIS...")
	findUCLeft(gisRows, gisColIndex)
	log.Println("GIS..OK!")
	findUCLeftByGIS2()
	log.Print("Right...OK")

	log.Println("Matching... OK!")
	return nil
}

func readIcbOrd()(error, col_index_t, [][]string){
	sheet := data.conf.sheets["IO_SHEET"]
	colIndex := newColIndex()
	rows := data.ioXlsx.GetRows(sheet)
	if len(rows) <= 1 {
		return errors.New("Can not find '" + sheet + "' sheet"), colIndex, nil
	}

	data.ioHeader = rows[0]
	
	for index, name := range data.ioHeader {
		if name == "SO No." {
			colIndex.soNo = index
		}else if name == "WBS" {
			colIndex.wbs = index
		}else if name == "DB SO No." {
			colIndex.dbSoNo = index
		}
	}

	if colIndex.soNo < 0 || colIndex.wbs < 0 || colIndex.dbSoNo < 0{
		return errors.New("Can not find SO No./WBS/DB SO No. columns"), colIndex, nil
	}

	rows = rows[1:]
	newRows := make([][]string, 0)
	for _, row := range rows {
		if util.IsEmptyRow(row) {
			continue
		}
		row[colIndex.soNo] = strings.TrimSpace(row[colIndex.soNo])
		row[colIndex.wbs] = parseWbsNoValue(strings.TrimSpace(row[colIndex.wbs]))
		row[colIndex.dbSoNo] = strings.TrimSpace(row[colIndex.dbSoNo])

		newRows = append(newRows, row)
	}
	
	return nil, colIndex, newRows
}

func readGIS()(error, col_index_t, [][]string){
	sheet := data.conf.sheets["GIS_SHEET"]
	colIndex := newColIndex()
	headerLineNo := 5
	rows := data.gisXlsx.GetRows(sheet)
	if len(rows) <= headerLineNo {
		return errors.New("Can not find '" + sheet + "' sheet header"), colIndex, nil
	}

	data.gisHeader = rows[headerLineNo - 1]


	for index, name := range data.gisHeader {
		name = strings.TrimSpace(name)
		if name == "SAP Order No.                   Segment  SO Number" {
			colIndex.dbSoNo = index
		}else if name == "CCM--WBS No." {
			colIndex.wbs = index
		}else if name == "SAP Order No.                 Operation  SO number" {
			colIndex.soNo = index
		}else if name == "Product" {
			colIndex.product = index
		}else if name == "Project Name" {
			colIndex.projectName = index
		}else if name == "Customer No." {
			colIndex.customerNo = index
		}else if name == "Customer Name" {
			colIndex.customerName = index
		}else if name == "Contract No." {
			colIndex.contractNo = index
		}
	}

	if colIndex.soNo < 0 || colIndex.wbs < 0 || colIndex.dbSoNo < 0{
		return errors.New("Can not find SO No./WBS/DB SO No. columns"), colIndex, nil
	}

	rows = rows[headerLineNo:]
	var newRows [][]string
	for _, row := range rows {
		if util.IsEmptyRow(row) {
			continue
		}
		row[colIndex.soNo] = strings.TrimSpace(row[colIndex.soNo])
		row[colIndex.wbs] = parseWbsNoValue(strings.TrimSpace(row[colIndex.wbs]))
		row[colIndex.dbSoNo] = strings.TrimSpace(row[colIndex.dbSoNo])

		newRows = append(newRows, row)
	}
	
	return nil, colIndex, newRows
}

func findUCRightByGIS2()error{
	log.Println("GIS2...")
	for _, item := range data.conf.gis2Sheets{
		err,colIdex,rows := readGIS2Sheet(item)
		if(err != nil){
			continue
		}

		err = findUCRight(rows, colIdex)
		if(err != nil) {
			continue
		}
		log.Println(item[0] + "..OK!")
	}
	log.Println("GIS2..OK!")
	return nil
}

func findUCLeftByGIS2()error{
	log.Println("GIS2...")
	for _, item := range data.conf.gis2Sheets{
		err,colIdex,rows := readGIS2Sheet(item)
		if(err != nil){
			continue
		}

		err = findUCLeft(rows, colIdex)
		if(err != nil) {
			continue
		}
		log.Println(item[0] + "..OK!")
	}
	log.Println("GIS2..OK!")
	return nil
}

func readGIS2Sheet(item [2]string)(error, col_index_t, [][]string){
	sheet, headerLineNoStr := item[0], item[1]
		log.Println(sheet + "...")
		headerLineNo, err := strconv.Atoi(headerLineNoStr)
		colIndex := newColIndex()
		if(err != nil) {
			log.Printf("Header Line No error(%s)\n", headerLineNoStr)
			return err,colIndex,nil
		}

		rows := data.gis2Xlsx.GetRows(sheet)
		if len(rows) <= headerLineNo {
			log.Printf("Can not find '%s' sheet\n", sheet)
			return err,colIndex,nil
		}

		for index, name := range rows[headerLineNo - 1] {
			if name == "" {
				continue
			}
			name = strings.ToLower(strings.TrimSpace(name))

			if colIndex.soNo == -1 && strings.Contains(name, "operation") {
				colIndex.soNo = index
			} else if colIndex.wbs == -1 && strings.Contains(name, "ccm--wbs") {
				colIndex.wbs = index
			} else if colIndex.dbSoNo == -1 && strings.Contains(name, "segment") {
				colIndex.dbSoNo = index
			} else if colIndex.product == -1 && strings.TrimSpace(name) == "product" {
				colIndex.product = index
			} else if colIndex.product == -1 && strings.TrimSpace(name) == "project name" {
				colIndex.projectName = index
			} else if colIndex.product == -1 && strings.TrimSpace(name) == "customer no." {
				colIndex.customerNo = index
			} else if colIndex.product == -1 && strings.TrimSpace(name) == "customer name" {
				colIndex.customerName = index
			} else if colIndex.product == -1 && strings.TrimSpace(name) == "contract no" {
				colIndex.contractNo = index
			}
		}

		if colIndex.soNo == -1 {
			return errors.New("Can not find Operation No column in '" + sheet + "' sheet"),colIndex,nil
		}
		if colIndex.dbSoNo == -1 {
			return errors.New("Can not find Segment No column in '" + sheet + "' sheet"),colIndex,nil
		}

		rows = rows[headerLineNo:]
		var newRows [][]string
		for _, row := range rows {
			if util.IsEmptyRow(row) {
				continue
			}
			row[colIndex.soNo] = strings.TrimSpace(row[colIndex.soNo])
			row[colIndex.dbSoNo] = strings.TrimSpace(row[colIndex.dbSoNo])
			if colIndex.wbs >= 0 {
				row[colIndex.wbs] = parseWbsNoValue(strings.TrimSpace(row[colIndex.wbs]))
			}
			newRows = append(newRows, row)
		}

		return nil,colIndex,newRows
}

func findUCRight(rows [][]string, colIndex col_index_t)error{
	for leftIndex, left := range data.ucLeftData{
		ucObjVal := data.ucData[left.idx][data.ucObjColIdx]
		for _, row := range rows{
			wbsVal := ""
			if(colIndex.wbs >= 0){
				wbsVal = row[colIndex.wbs]
			}
			dbSoNoVal := strings.TrimSpace(row[colIndex.dbSoNo])
			if(wbsVal != "" && ucObjVal == wbsVal || ucObjVal == dbSoNoVal){
				if colIndex.product > 0{
					data.ucLeftData[leftIndex].product = row[colIndex.product]
				}
				if colIndex.projectName > 0 {
					data.ucLeftData[leftIndex].projectName = row[colIndex.projectName]
				}
				if colIndex.customerNo > 0 {
					data.ucLeftData[leftIndex].customerNo = row[colIndex.customerNo]
				}
				if colIndex.customerName > 0 {
					data.ucLeftData[leftIndex].customerName = row[colIndex.customerName]
				}
				if colIndex.contractNo > 0 {
					data.ucLeftData[leftIndex].contractNo = row[colIndex.contractNo]
				}

				soNoVal := strings.TrimSpace(row[colIndex.soNo])
				isFindRight := false
				for rightIndx, right := range data.ucRightData{
					if right.leftIdx != DB_IDX_NO_RELATION && right.leftIdx != DB_IDX_NO_DATA{
						continue
					}

					if soNoVal == data.ucData[right.idx][data.ucObjColIdx]{
						data.ucLeftData[leftIndex].rightIdxs = append(data.ucLeftData[leftIndex].rightIdxs, right.idx)
						data.ucRightData[rightIndx].leftIdx = left.idx

						isFindRight = true
					}
				}

				if(!isFindRight && soNoVal != ""){
					isAdded := false
					for _,addedSoNo := range(data.ucLeftData[leftIndex].noRightSoNoList){
						if addedSoNo == soNoVal{
							isAdded = true
							break;
						}
					}
					if !isAdded {
						data.ucLeftData[leftIndex].noRightSoNoList = append(data.ucLeftData[leftIndex].noRightSoNoList, soNoVal)
					}
				}
			}
		}
	}
	return nil
}

func findUCLeft(rows [][]string, colIndex col_index_t)error{
	for rightIndex, right := range data.ucRightData{
		if right.leftIdx != DB_IDX_NO_RELATION && right.leftIdx != DB_IDX_NO_DATA{
			continue
		}

		ucObjVal := data.ucData[right.idx][data.ucObjColIdx]
		for _, row := range rows{
			soNoVal := strings.TrimSpace(row[colIndex.soNo])
			if(ucObjVal == soNoVal){
				if colIndex.product > 0{
					data.ucRightData[rightIndex].product = row[colIndex.product]
				}
				if colIndex.projectName > 0 {
					data.ucRightData[rightIndex].projectName = row[colIndex.projectName]
				}
				if colIndex.customerNo > 0 {
					data.ucRightData[rightIndex].customerNo = row[colIndex.customerNo]
				}
				if colIndex.customerName > 0 {
					data.ucRightData[rightIndex].customerName = row[colIndex.customerName]
				}
				if colIndex.contractNo > 0 {
					data.ucRightData[rightIndex].contractNo = row[colIndex.contractNo]
				}

				wbsVal := ""
				if(colIndex.wbs >= 0){
					wbsVal = row[colIndex.wbs]
				}
				if wbsVal != ""{
					data.ucRightData[rightIndex].dbSoNo = wbsVal
				}else{
					dbSoNoVal := strings.TrimSpace(row[colIndex.dbSoNo])
					data.ucRightData[rightIndex].dbSoNo = dbSoNoVal
				}
				data.ucRightData[rightIndex].leftIdx = DB_IDX_NO_DATA
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
