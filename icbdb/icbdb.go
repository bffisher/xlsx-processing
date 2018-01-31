package icbdb

import (
	"errors"
	"log"
	"strconv"
	"strings"
	"xlsx-processing/util"

	"github.com/360EntSecGroup-Skylar/excelize"
)

//db is some rows of OD_COPA rows
//idx: db index of odCopaRows, cnt: found count by icb
type db_t struct {
	idx, cnt int
}

//icb is some rows of OD_COPA rows, save icb and relation of icb and db
//idx: icb index of odCopaRows
//db: corresponding to icb
type icb_db_t struct {
	idx int
	db  db_t
}

type data_t struct {
	path         string
	conf         *config_t
	odXlsx       *excelize.File
	odCopaRows   [][]string
	odCopaHeader map[string]int
	icbDbList    []icb_db_t
	dbList       []db_t
}

const _DB_IDX_EMPTY = -1
const _DB_IDX_NIL = -9
const _ICB_IDX_NIL = -9

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
	appendUnmatchedDb(data)
	log.Println("OK!")

	log.Print("Outputing... ")
	output(data)
	log.Println("OK!")
	return nil
}

func splitIcbDb(data *data_t) {
	len := len(data.odCopaRows)
	icbDbList := make([]icb_db_t, 0, len/3*2)
	dbList := make([]db_t, 0, len/2)

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
			icbDbList = append(icbDbList, icb_db_t{index, db_t{_DB_IDX_NIL, 0}})
		} else {
			//DB
			dbList = append(dbList, db_t{index, 0})
		}
	}
	data.icbDbList, data.dbList = icbDbList, dbList
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
	count, isLeft := 0, false
	for icbIndex, icbDb := range data.icbDbList {
		if icbDb.db.idx != _DB_IDX_NIL {
			continue
		}
		icbSoNo := strings.TrimSpace(data.odCopaRows[icbDb.idx][coapSoNoColIdx])
		index := -1
		for _, odIcbOrdRow := range rows {
			if icbSoNo == strings.TrimSpace(odIcbOrdRow[soNoColIdx]) {
				wbs := ""
				if wbsColIdx >= 0 {
					wbs = strings.TrimSpace(odIcbOrdRow[wbsColIdx])
				}
				if wbs != "" {
					index = matchODCopaWBS(data, wbs, icbDb.idx)
				} else {
					dbSoNo := strings.TrimSpace(odIcbOrdRow[dbSoNoColIdx])
					index = matcODCopaSoNo(data, dbSoNo, icbDb.idx)
				}
				if index >= 0 {
					data.dbList[index].cnt++
					data.icbDbList[icbIndex].db = data.dbList[index]
				} else {
					data.icbDbList[icbIndex].db = db_t{_DB_IDX_EMPTY, 0}
				}
				count++
				break
			}
		}
		if !isLeft && data.icbDbList[icbIndex].db.idx == _DB_IDX_NIL {
			isLeft = true
		}
	}
	log.Println("Found:", count)
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

func appendUnmatchedDb(data *data_t) {
	for _, db := range data.dbList {
		if db.cnt == 0 {
			// have not a corresponding icb
			data.icbDbList = append(data.icbDbList, icb_db_t{_ICB_IDX_NIL, db})
		}
	}
}
