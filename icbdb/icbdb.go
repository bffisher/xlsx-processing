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
	icbSoNo string
	province, classfication string
}

type data_t struct {
	conf         *config_t
	odXlsx, gisXlsx, mcXlsx, gis19Xlsx,vi19Xlsx *excelize.File
	odCopaRows   [][]string
	odCopaHeader map[string]int
	icbList      []icb_t
	dbList       []db_t
	vi19sheetInfo [][2] string
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

func Exec() error {
	data := &data_t{}

	log.Print("Reading... ")
	conf, err := readConfig(util.Env().ConfigFullName)
	if err != nil {
		return err
	}
	data.conf = conf

	data.vi19sheetInfo = append(data.vi19sheetInfo, [2]string{"Order List FY2019", "4"})
	data.vi19sheetInfo = append(data.vi19sheetInfo, [2]string{"Orderlist_VI Parts FY2019", "6"})
	data.vi19sheetInfo = append(data.vi19sheetInfo, [2]string{"Order List FY2018", "4"})
	data.vi19sheetInfo = append(data.vi19sheetInfo, [2]string{"Orderlist_VI Parts FY2018", "6"})


	odXlsx, err := excelize.OpenFile(util.Env().FilePath + conf.files["OD"])
	if err != nil {
		return err
	}
	gisXlsx, err := excelize.OpenFile(util.Env().FilePath + conf.files["GIS"])
	if err != nil {
		return err
	}
	mcXlsx, err := excelize.OpenFile(util.Env().FilePath + conf.files["MC"])
	if err != nil {
		return err
	}
	gis19Xlsx, err := excelize.OpenFile(util.Env().FilePath + conf.files["GIS19"])
	if err != nil {
		return err
	}
	vi19Xlsx, err := excelize.OpenFile(util.Env().FilePath + conf.files["VI19"])
	if err != nil {
		return err
	}
	data.odXlsx = odXlsx
	data.gisXlsx = gisXlsx
	data.mcXlsx = mcXlsx
	data.gis19Xlsx = gis19Xlsx
	data.vi19Xlsx = vi19Xlsx

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
			dbList = append(dbList, db_t{index, make([]int, 0), "", "", ""})
		}
	}
	data.icbList, data.dbList = icbList, dbList
}

func resolveIcbDbRelation(data *data_t) error {
	err := findDbInDbListByODIcbOrd(data)
	if err != nil {
		return err
	}

	err =  findDbInDblistByOther(data, data.gisXlsx, data.conf.gisSheets)
	if err != nil {
		return err
	}

	err = findDbInDblistByOther(data, data.mcXlsx, data.conf.mcSheets)
	if err != nil {
		return err
	}

	err = findDbInDblistByGis19(data)
	if err != nil {
		return err
	}

	err = findDbInDblistByVi19(data)
	if err != nil {
		return err
	}
	return nil
}

func findDbInDbListByODIcbOrd(data *data_t) error {
	sheet := data.conf.sheets["OD_ICB_ORD"]
	log.Println(sheet)
	odIcbOrdRows := data.odXlsx.GetRows(data.conf.sheets["OD_ICB_ORD"])
	if len(odIcbOrdRows) < 2 {
		return errors.New("Can not find '" + sheet + "' sheet")
	}

	icbDBHeader, err := getODIcbOrdHeader(odIcbOrdRows[0], data.conf)
	if err != nil {
		return err
	}

	soNoColIdx := icbDBHeader["OD_ICB_ORD_SN"]
	wbsColIdx := icbDBHeader["OD_ICB_ORD_WBS"]
	dbSoNoColIdx := icbDBHeader["OD_ICB_ORD_DB_SN"]
	findDbInDbList(data, odIcbOrdRows, soNoColIdx, wbsColIdx, dbSoNoColIdx)
	return  nil
}

func findDbInDblistByOther(data *data_t, xlsx *excelize.File, sheets [][3]string) error {

	for _, item := range sheets {
		sheetData,err:= getOtherSheetData(data, xlsx, item)
		if err!=nil{
			return err
		}

		findDbInDbList(data, sheetData.rows[sheetData.headerIdx:], sheetData.soNoColIdx, sheetData.wbsColIdx, sheetData.dbSoNoColIdx)
	}
	return nil
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

func findDbInDblistByGis19(data *data_t)error{
	log.Printf("[%s] %s", data.gis19Xlsx.Path, "Sheet1")
	err,colIndex, rows := util.ReadGIS19(data.gis19Xlsx, "Sheet1")
	if(err!=nil){
		return err
	}
	findDbInDbList(data, rows, colIndex.SoNo, colIndex.Wbs, colIndex.DbSoNo)
	return nil
}

func findDbInDblistByVi19(data *data_t)error{
	for _,sheetInfo := range(data.vi19sheetInfo){
		log.Printf("[%s] %s", data.vi19Xlsx.Path, sheetInfo[0])
		err,colIndex, rows := util.ReadVi19(data.vi19Xlsx, sheetInfo[0], sheetInfo[1])
		if(err!=nil){
			return err
		}
		findDbInDbList(data, rows, colIndex.SoNo, colIndex.Wbs, colIndex.DbSoNo)
	}
	return nil
}

func findDbInDbList(data *data_t, rows [][]string, soNoColIdx, wbsColIdx, dbSoNoColIdx int) {
	coapSoNoColIdx := data.odCopaHeader["OD_COPA_SO"]
	coapWbsColIdx := data.odCopaHeader["OD_COPA_WBS"]
	rcount, dcount:= 0, 0
	for dbIndex, db := range data.dbList {
		if len(db.icbIdxs) > 0{
			continue
		}
		dbWbs := strings.TrimSpace(data.odCopaRows[db.idx][coapWbsColIdx])
		dbSoNo := strings.TrimSpace(data.odCopaRows[db.idx][coapSoNoColIdx])
		for _, row := range rows {
			rdbWbs := ""
			if wbsColIdx >= 0 {
				rdbWbs = strings.TrimSpace(row[wbsColIdx])
			}

			ricbSoNo := strings.TrimSpace(row[soNoColIdx])
			rdbSoNo := strings.TrimSpace(row[dbSoNoColIdx])
			if rdbWbs != "" && dbWbs != "" && strings.Contains(rdbWbs, dbWbs[0:17]) || dbSoNo == rdbSoNo {
				for icbIndex, icb := range data.icbList {
					if icb.dbIdx != _DB_IDX_NO_RELATION && icb.dbIdx != _DB_IDX_NO_DATA{
						continue
					}
					icbSoNo := strings.TrimSpace(data.odCopaRows[icb.idx][coapSoNoColIdx])				
					if ricbSoNo == icbSoNo {
						data.dbList[dbIndex].icbIdxs = append(data.dbList[dbIndex].icbIdxs, icb.idx)
						data.icbList[icbIndex].dbIdx = db.idx
						dcount++
					} 
				}
				//event if found in relation table but conuld not find in icb list
				if len(data.dbList[dbIndex].icbIdxs) == 0{
					data.dbList[dbIndex].icbSoNo = ricbSoNo
				}
				rcount++
			}
		}
	}
	log.Printf("Found:%d(%d)", rcount, dcount)
	for icbIndex, icb := range data.icbList {
		if icb.dbIdx == _DB_IDX_NO_RELATION{
			for _, row := range rows {
				ricbSoNo := strings.TrimSpace(row[soNoColIdx])
				icbSoNo := strings.TrimSpace(data.odCopaRows[icb.idx][coapSoNoColIdx])
				if ricbSoNo == icbSoNo {
					data.icbList[icbIndex].dbIdx = _DB_IDX_NO_DATA
					if wbsColIdx >= 0 {
						data.icbList[icbIndex].dbWbs = strings.TrimSpace(row[wbsColIdx])
					}
					data.icbList[icbIndex].dbSoNo = strings.TrimSpace(row[dbSoNoColIdx])
				}
			}
		}
	}
}

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

	log.Printf("[%s] %s", data.gis19Xlsx.Path, "Sheet1")
	err,gis19ColIndex, gis19Rows := util.ReadGIS19(data.gis19Xlsx, "Sheet1")
	if(err!=nil){
		return err
	}
	gis19SheetData := other_sheet_data_t{}
	gis19SheetData.classficationColIdx = gis19ColIndex.Classfication
	gis19SheetData.soNoColIdx = gis19ColIndex.SoNo
	gis19SheetData.wbsColIdx = gis19ColIndex.Wbs
	gis19SheetData.dbSoNoColIdx = gis19ColIndex.DbSoNo
	for index, db:= range data.dbList{
		if db.classfication == "" && gis19ColIndex.Classfication >= 0{
			data.dbList[index].classfication = findColValInOterSheetByDb(data, &db, gis19Rows,	&gis19SheetData, gis19SheetData.classficationColIdx, "Y")
		}
	}
	for index, icb := range data.icbList {
		if icb.dbIdx == _DB_IDX_NO_DATA || icb.dbIdx == _DB_IDX_NO_RELATION{
			if icb.classfication == "" && gis19SheetData.classficationColIdx >= 0{
				data.icbList[index].classfication = findColValInOterSheetByIcb(data, &icb, gis19Rows, &gis19SheetData, gis19SheetData.classficationColIdx, "N")
			}
		}
	}

	for _,sheetInfo := range(data.vi19sheetInfo){
		log.Printf("[%s] %s", data.vi19Xlsx.Path, sheetInfo[0])
		err,colIndex, rows := util.ReadVi19(data.vi19Xlsx, sheetInfo[0], sheetInfo[1])
		if(err!=nil){
			return err
		}
		sheetData := other_sheet_data_t{}
		sheetData.classficationColIdx = colIndex.Classfication
		sheetData.soNoColIdx = colIndex.SoNo
		sheetData.wbsColIdx = colIndex.Wbs
		sheetData.dbSoNoColIdx = colIndex.DbSoNo
		for index, db:= range data.dbList{
			if db.classfication == "" && sheetData.classficationColIdx >= 0{
				data.dbList[index].classfication = findColValInOterSheetByDb(data, &db, rows,	&sheetData, sheetData.classficationColIdx, "Y")
			}
		}
		for index, icb := range data.icbList {
			if icb.dbIdx == _DB_IDX_NO_DATA || icb.dbIdx == _DB_IDX_NO_RELATION{
				if icb.classfication == "" && sheetData.classficationColIdx >= 0{
					data.icbList[index].classfication = findColValInOterSheetByIcb(data, &icb, rows, &sheetData, sheetData.classficationColIdx, "N")
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