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
	wbs, soNo  string
}

//db is some rows of OD_COPA rows
//idx: index of odCopaRows
//icbIdxs: indexs of odCopaRows, that belong to icb
type db_t struct {
	idx     int
	icbIdxs []int
}

type data_t struct {
	path         string
	conf         *config_t
	odXlsx       *excelize.File
	odCopaRows   [][]string
	odCopaHeader map[string]int
	icbList      []icb_t
	dbList       []db_t
}

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
	data.odXlsx = odXlsx

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
	soNoColIdx := data.odCopaHeader["OD_COPA_SO"]
	tradPartnColIdx := data.odCopaHeader["OD_COPA_TP"]
	for index, row := range data.odCopaRows {
		if wbs := row[wbsColIdx]; wbs != "" {
			_, wbs = util.SplitCodeName(wbs)
			if strings.Contains(strings.ToUpper(wbs), "POC") {
				continue
			}
		}

		soNo := row[soNoColIdx]
		tradPartn, _ := util.SplitCodeName(row[tradPartnColIdx])
		if soNo != "" && tradPartn == "004611" {
			//ICB
			icbList = append(icbList, icb_t{index, _DB_IDX_NO_RELATION, "", ""})
		} else {
			//DB
			dbList = append(dbList, db_t{index, make([]int, 0)})
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

	isLeft, err = findDbInDblistByGis(data)
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

func findDbInDblistByGis(data *data_t) (bool, error) {
	isLeft := false
	xlsx, err := excelize.OpenFile(data.path + data.conf.files["GIS"])
	if err != nil {
		return false, err
	}
	for _, item := range data.conf.gisSheets {
		sheet, headerIdxStr := item[0], item[1]
		log.Println(item[0])
		headerIdx, err := strconv.Atoi(headerIdxStr)
		if err != nil {
			return false, err
		}
		rows := xlsx.GetRows(sheet)
		if len(rows) <= headerIdx {
			return false, errors.New("Can not find '" + sheet + "' sheet")
		}
		soNoColIdx, wbsColIdx, dbSoNoColIdx := -1, -1, -1
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
			}
		}

		if soNoColIdx == -1 {
			return false, errors.New("Can not find Operation No column in '" + sheet + "' sheet")
		}
		if dbSoNoColIdx == -1 {
			return false, errors.New("Can not find Segment No column in '" + sheet + "' sheet")
		}

		isLeft = findDbInDbList(data, rows[headerIdx:], soNoColIdx, wbsColIdx, dbSoNoColIdx)
		if !isLeft {
			break
		}
	}
	return isLeft, nil
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
				wbs, dbSoNo := "", ""
				if wbsColIdx >= 0 {
					wbs = strings.TrimSpace(row[wbsColIdx])
				}
				if wbs != "" {
					idxInDbList = matchODCopaWBS(data, wbs, icb.idx)
				} else {
					dbSoNo = strings.TrimSpace(row[dbSoNoColIdx])
					idxInDbList = matcODCopaSoNo(data, dbSoNo, icb.idx)
				}
				if idxInDbList >= 0 {
					data.dbList[idxInDbList].icbIdxs = append(data.dbList[idxInDbList].icbIdxs, icb.idx)
					data.icbList[index].dbIdx = data.dbList[idxInDbList].idx
				} else {
					data.icbList[index].dbIdx = _DB_IDX_NO_DATA
					data.icbList[index].wbs = wbs
					data.icbList[index].soNo = dbSoNo
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
	productHierarchyColIdx := data.odCopaHeader["OD_COPA_PH"]

	for index, db := range data.dbList {
		dbWbs := strings.TrimSpace(data.odCopaRows[db.idx][wbsColIdx])
		if dbWbs != "" && wbs[0:17] == dbWbs[0:17] {
			icbProductHierarchy := strings.TrimSpace(data.odCopaRows[rowIdx][productHierarchyColIdx])
			dbProductHierarchy := strings.TrimSpace(data.odCopaRows[db.idx][productHierarchyColIdx])
			if matchProductHierarchy(data, icbProductHierarchy, dbProductHierarchy) {
				return index
			}
		}
	}
	return -1
}

func matcODCopaSoNo(data *data_t, dbSoNo string, rowIdx int) int {
	soNoColIdx := data.odCopaHeader["OD_COPA_SO"]
	productHierarchyColIdx := data.odCopaHeader["OD_COPA_PH"]

	for index, db := range data.dbList {
		if dbSoNo == strings.TrimSpace(data.odCopaRows[db.idx][soNoColIdx]) {
			icbProductHierarchy := strings.TrimSpace(data.odCopaRows[rowIdx][productHierarchyColIdx])
			dbProductHierarchy := strings.TrimSpace(data.odCopaRows[db.idx][productHierarchyColIdx])
			if matchProductHierarchy(data, icbProductHierarchy, dbProductHierarchy) {
				return index
			}
		}
	}
	return -1
}

func matchProductHierarchy(data *data_t, val1, val2 string) bool {
	_, name1 := util.SplitCodeName(val1)
	_, name2 := util.SplitCodeName(val2)

	return data.conf.products[name1] == data.conf.products[name2]
}
