package icbdb

import (
	"errors"
	"log"
	"strconv"
	"strings"
	"xlsx-processing/util"

	"github.com/360EntSecGroup-Skylar/excelize"
)

//icb is some rows of OD_COPA rows
//idx: index of odCopaRows
//dbIdx: index of odCopaRows, that belong to db
type icb_t struct {
	idx, dbIdx int
	dbWbs, dbSoNo  string
	province, classfication string
}

//db is some rows of OD_COPA rows
//idx: index of odCopaRows
//icbIdxs: indexs of odCopaRows, that belong to icb
type db_t struct {
	idx     int
	icbIdxs []int
	province, classfication string
}

type data_t struct {
	path         string
	conf         *config_t
	odXlsx, gisXlsx, mcXlsx       *excelize.File
	odCopaRows   [][]string
	odCopaHeader map[string]int
	icbList      []icb_t
	dbList       []db_t
}

type other_sheet_data_t struct{
	rows [][]string
	headerIdx int
	soNoColIdx, wbsColIdx, dbSoNoColIdx, provinceColIdx, partnerColIdx, classficationColIdx int
}

const _TR_4611 = "004611"
const _PC_P8251 = "P8251"
const _PC_P8211 = "P8211"

const _DB_IDX_NO_RELATION = -1
const _DB_IDX_NO_DATA = -2

func Exec(confFile string) error {
	data := &data_t{}

	log.Print("Reading... ")
	conf, err := readConfig(data.path + confFile)
	if err != nil {
		return err
	}
	data.conf = conf

	odXlsx, err := excelize.OpenFile(data.path + conf.files["OD"])
	if err != nil {
		return err
	}
	gisXlsx, err := excelize.OpenFile(data.path + conf.files["GIS"])
	if err != nil {
		return err
	}
	mcXlsx, err := excelize.OpenFile(data.path + conf.files["MC"])
	if err != nil {
		return err
	}
	data.odXlsx = odXlsx
	data.gisXlsx = gisXlsx
	data.mcXlsx = mcXlsx

	odCopaRows := odXlsx.GetRows(conf.sheets["OD_COPA"])
	if len(odCopaRows) < 2 {
		return errors.New("Can not find 'COPA original data' sheet")
	}

	odCopaHeader, err := getODCopaHeader(odCopaRows[0], conf)
	if err != nil {
		return err
	}
	data.odCopaHeader = odCopaHeader
	data.odCopaRows = odCopaRows[1:]
	log.Println("OK!")

	log.Print("Matching... ")
	splitIcbDb(data)
	err = resolveIcbDbRelation(data)
	if err != nil {
		return err
	}

	log.Print("Finding province and classfication...")
	findProvinceAndClassfication(data)

	log.Println("OK!")

	log.Print("Outputing... ")
	output(data)
	log.Println("OK!")
	return nil
}

func splitIcbDb(data *data_t) {
	len := len(data.odCopaRows)
	icbList := make([]icb_t, 0, len/2)
	dbList := make([]db_t, 0, len/3*2)

	wbsColIdx := data.odCopaHeader["OD_COPA_WBS"]
	// soNoColIdx := data.odCopaHeader["OD_COPA_SO"]
	tradPartnColIdx := data.odCopaHeader["OD_COPA_TP"]
	for index, row := range data.odCopaRows {
		if wbs := row[wbsColIdx]; wbs != "" {
			if strings.Contains(strings.ToUpper(wbs), "POC") {
				continue
			}
		}

		// soNo := row[soNoColIdx]
		tradPartn, _ := util.SplitCodeName(row[tradPartnColIdx])
		if /*soNo != "" &&*/ tradPartn == _TR_4611 {
			//ICB
			icbList = append(icbList, icb_t{index, _DB_IDX_NO_RELATION, "", "", "", ""})
		} else {
			//DB
			dbList = append(dbList, db_t{index, make([]int, 0), "", ""})
		}
	}
	data.icbList, data.dbList = icbList, dbList
}

func resolveIcbDbRelation(data *data_t) error {
	isLeft, err := findDbInDbListByODIcbOrd(data)
	if err != nil {
		return err
	}
	if !isLeft {
		return nil
	}

	isLeft, err =  findDbInDblistByOther(data, data.gisXlsx, data.conf.gisSheets)
	if err != nil {
		return err
	}

	isLeft, err = findDbInDblistByOther(data, data.mcXlsx, data.conf.mcSheets)
	if err != nil {
		return err
	}
	return nil
}

func findDbInDbListByODIcbOrd(data *data_t) (bool, error) {
	sheet := data.conf.sheets["OD_ICB_ORD"]
	log.Println(sheet)
	odIcbOrdRows := data.odXlsx.GetRows(data.conf.sheets["OD_ICB_ORD"])
	if len(odIcbOrdRows) < 2 {
		return false, errors.New("Can not find '" + sheet + "' sheet")
	}

	icbDBHeader, err := getODIcbOrdHeader(odIcbOrdRows[0], data.conf)
	if err != nil {
		return false, err
	}

	soNoColIdx := icbDBHeader["OD_ICB_ORD_SN"]
	wbsColIdx := icbDBHeader["OD_ICB_ORD_WBS"]
	dbSoNoColIdx := icbDBHeader["OD_ICB_ORD_DB_SN"]
	return findDbInDbList(data, odIcbOrdRows, soNoColIdx, wbsColIdx, dbSoNoColIdx), nil
}

func findDbInDblistByOther(data *data_t, xlsx *excelize.File, sheets [][3]string) (bool, error) {
	isLeft := false
	
	for _, item := range sheets {
		sheetData,err:= getOtherSheetData(data, xlsx, item)
		if err!=nil{
			return false, err
		}

		isLeft = findDbInDbList(data, sheetData.rows[sheetData.headerIdx:], sheetData.soNoColIdx, sheetData.wbsColIdx, sheetData.dbSoNoColIdx)
		if !isLeft {
			break
		}
	}
	return isLeft, nil
}

func getOtherSheetData(data *data_t, xlsx *excelize.File, sheetInfo [3]string) (other_sheet_data_t,error){
	result := other_sheet_data_t{};
	sheet, headerIdxStr := sheetInfo[0], sheetInfo[1]
	log.Printf("[%s] %s", xlsx.Path, sheet)
	headerIdx, err := strconv.Atoi(headerIdxStr)
	if err != nil {
		return result, err;
	}
	rows := xlsx.GetRows(sheet)
	if len(rows) <= headerIdx {
		return result, errors.New("Can not find '" + sheet + "' sheet")
	}
	soNoColIdx, wbsColIdx, dbSoNoColIdx, provinceColIdx, partnerColIdx, classficationColIdx := -1, -1, -1, -1,-1,-1
	for index, name := range rows[headerIdx-1] {
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
		}else if provinceColIdx == -1 && strings.Contains(name, "province"){
			provinceColIdx = index
		}else if partnerColIdx == -1 && strings.Contains(name, "partner"){
			partnerColIdx = index
		}else if classficationColIdx == -1 && strings.Contains(name, "classfication"){
			classficationColIdx = index
		}
	}

	if soNoColIdx == -1 {
		return result, errors.New("Can not find Operation No column in '" + sheet + "' sheet")
	}
	if dbSoNoColIdx == -1 {
		return result, errors.New("Can not find Segment No column in '" + sheet + "' sheet")
	}

	result.rows = rows
	result.headerIdx = headerIdx
	result.soNoColIdx = soNoColIdx
	result.wbsColIdx = wbsColIdx
	result.dbSoNoColIdx = dbSoNoColIdx
	result.provinceColIdx = provinceColIdx
	result.partnerColIdx = partnerColIdx
	result.classficationColIdx = classficationColIdx
	return result, nil
}

func findDbInDbList(data *data_t, rows [][]string, soNoColIdx, wbsColIdx, dbSoNoColIdx int) bool {
	coapSoNoColIdx := data.odCopaHeader["OD_COPA_SO"]
	rcount, dcount, isLeft := 0, 0, false
	for index, icb := range data.icbList {
		if icb.dbIdx != _DB_IDX_NO_RELATION {
			continue
		}
		icbSoNo := strings.TrimSpace(data.odCopaRows[icb.idx][coapSoNoColIdx])
		idxInDbList := -1
		for _, row := range rows {
			if icbSoNo == strings.TrimSpace(row[soNoColIdx]) {
				dbWbs, dbSoNo := "", ""
				if wbsColIdx >= 0 {
					dbWbs = strings.TrimSpace(row[wbsColIdx])
				}
				if dbWbs != "" {
					idxInDbList = matchODCopaWBS(data, dbWbs, icb.idx)
				} else {
					if dbSoNo = strings.TrimSpace(row[dbSoNoColIdx]); dbSoNo != ""{
						idxInDbList = matcODCopaSoNo(data, dbSoNo, icb.idx)
					}else{
						continue
					}
				}
				if idxInDbList >= 0 {
					data.dbList[idxInDbList].icbIdxs = append(data.dbList[idxInDbList].icbIdxs, icb.idx)
					data.icbList[index].dbIdx = data.dbList[idxInDbList].idx
				} else {
					data.icbList[index].dbIdx = _DB_IDX_NO_DATA
					data.icbList[index].dbWbs = dbWbs
					data.icbList[index].dbSoNo = dbSoNo
				}
				if idxInDbList >= 0 {
					dcount++
				}
				rcount++
				break
			}
		}
		if !isLeft && data.icbList[index].dbIdx == _DB_IDX_NO_RELATION {
			isLeft = true
		}
	}
	log.Printf("Found:%d(%d)", rcount, dcount)
	return isLeft
}

func matchODCopaWBS(data *data_t, wbs string, rowIdx int) int {
	wbsColIdx := data.odCopaHeader["OD_COPA_WBS"]
	// productHierarchyColIdx := data.odCopaHeader["OD_COPA_PH"]

	for index, db := range data.dbList {
		dbWbs := strings.TrimSpace(data.odCopaRows[db.idx][wbsColIdx])
		if dbWbs != "" && strings.Contains(wbs, dbWbs[0:17]) {
			// icbProductHierarchy := strings.TrimSpace(data.odCopaRows[rowIdx][productHierarchyColIdx])
			// dbProductHierarchy := strings.TrimSpace(data.odCopaRows[db.idx][productHierarchyColIdx])
			// if matchProductHierarchy(data, icbProductHierarchy, dbProductHierarchy) {
				return index
			// }
		}
	}
	return -1
}

func matcODCopaSoNo(data *data_t, dbSoNo string, rowIdx int) int {
	soNoColIdx := data.odCopaHeader["OD_COPA_SO"]
	// productHierarchyColIdx := data.odCopaHeader["OD_COPA_PH"]

	for index, db := range data.dbList {
		if dbSoNo == strings.TrimSpace(data.odCopaRows[db.idx][soNoColIdx]) {
			// icbProductHierarchy := strings.TrimSpace(data.odCopaRows[rowIdx][productHierarchyColIdx])
			// dbProductHierarchy := strings.TrimSpace(data.odCopaRows[db.idx][productHierarchyColIdx])
			// if matchProductHierarchy(data, icbProductHierarchy, dbProductHierarchy) {
				return index
			// }
		}
	}
	return -1
}

// func matchProductHierarchy(data *data_t, val1, val2 string) bool {
// 	_, name1 := util.SplitCodeName(val1)
// 	_, name2 := util.SplitCodeName(val2)

// 	return data.conf.products[name1] == data.conf.products[name2]
// }
func findProvinceAndClassfication(data *data_t)error{
	for _,sheetInfo:= range data.conf.gisSheets{
		
		sheetData,err:= getOtherSheetData(data, data.gisXlsx, sheetInfo)
		if err != nil{
			return err;
		}
		rows := sheetData.rows[sheetData.headerIdx:]
		for index, db:= range data.dbList{
			if sheetInfo[2] == "Y" && db.province == "" && sheetData.provinceColIdx >= 0{
				data.dbList[index].province = findColValInOterSheetByDb(data, &db, rows, &sheetData, sheetData.provinceColIdx, "Y")
			}

			if db.classfication == "" && sheetData.classficationColIdx >= 0{
				data.dbList[index].classfication = findColValInOterSheetByDb(data, &db, rows,	&sheetData, sheetData.classficationColIdx, "N")
			}
		}

		for index, icb := range data.icbList {
			if icb.dbIdx == _DB_IDX_NO_DATA || icb.dbIdx == _DB_IDX_NO_RELATION{
				if sheetInfo[2] == "Y" && icb.province == "" && sheetData.provinceColIdx >= 0{
					data.icbList[index].province = findColValInOterSheetByIcb(data, &icb, rows, &sheetData, sheetData.provinceColIdx, "Y")
				}
	
				if icb.classfication == "" && sheetData.classficationColIdx >= 0{
					data.icbList[index].classfication = findColValInOterSheetByIcb(data, &icb, rows,	&sheetData, sheetData.classficationColIdx, "N")
				}
			}
		}
	}
	
	for _,sheetInfo:= range data.conf.mcSheets{
			sheetData,err:= getOtherSheetData(data, data.mcXlsx, sheetInfo)
			if err != nil{
				return err;
			}
			rows := sheetData.rows[sheetData.headerIdx:]
			for index, db:= range data.dbList{
				if db.province == "" && sheetData.partnerColIdx >= 0{
					data.dbList[index].province = findColValInOterSheetByDb(data, &db, rows, &sheetData, sheetData.partnerColIdx, "Y")
				}

				if db.classfication == "" && sheetData.classficationColIdx >= 0{
					data.dbList[index].classfication = findColValInOterSheetByDb(data, &db, rows,	&sheetData, sheetData.classficationColIdx, "N")
				}
			} 

			for index, icb := range data.icbList {
				if icb.dbIdx == _DB_IDX_NO_DATA || icb.dbIdx == _DB_IDX_NO_RELATION{
	
					if icb.province == "" && sheetData.partnerColIdx >= 0{
						data.icbList[index].province = findColValInOterSheetByIcb(data, &icb, rows, &sheetData, sheetData.partnerColIdx, "Y")
					}

					if icb.classfication == "" && sheetData.classficationColIdx >= 0{
						data.icbList[index].classfication = findColValInOterSheetByIcb(data, &icb, rows,	&sheetData, sheetData.classficationColIdx, "N")
					}
				}
			}
	}
	return nil
}

func findColValInOterSheetByDb(data *data_t, db *db_t, rows [][]string, sheetData * other_sheet_data_t, valColIdx int, export string)string{
	val:= ""
	val = findColValInOtherSheet(data, db.idx, rows, sheetData.wbsColIdx, sheetData.dbSoNoColIdx, valColIdx, export)
	if val == ""{
		for _,icbIdx := range db.icbIdxs{
			val = findColValInOtherSheet(data, icbIdx, rows, sheetData.wbsColIdx, sheetData.soNoColIdx, valColIdx, export)
			if val != ""{
				break;
			}
		}
	}
	return val
}

func findColValInOterSheetByIcb(data *data_t, icb *icb_t, rows [][]string, sheetData * other_sheet_data_t, valColIdx int, export string)string{
	return findColValInOtherSheet(data, icb.idx, rows, sheetData.wbsColIdx, sheetData.dbSoNoColIdx, valColIdx, export)
}

func findColValInOtherSheet(data *data_t, idx int, rows [][]string, wbsColIdx, soNoColIdx, valColIdx int, export string)string{	
	wbs:= strings.TrimSpace(data.odCopaRows[idx][data.odCopaHeader["OD_COPA_WBS"]])
	soNo:= strings.TrimSpace(data.odCopaRows[idx][data.odCopaHeader["OD_COPA_SO"]])

	for _, row := range rows{
		if ex:= data.odCopaRows[idx][data.odCopaHeader["OD_COPA_EX"]];ex == export{
			if wbsColIdx >= 0 && wbs != "" && strings.Contains(strings.TrimSpace(row[wbsColIdx]), wbs[0:17]) {
				// log.Print("wbs",valColIdx)
				return row[valColIdx]
			}

			if soNoColIdx >= 0 && soNo != "" && soNo == strings.TrimSpace(row[soNoColIdx]){
				// log.Print("soNo",valColIdx)
				return row[valColIdx]
			}
		}	
	} 
	return  ""
}