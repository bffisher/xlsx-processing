package icbdb

import (
	"errors"
	"log"
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
	conf                     *config_t
	odXlsx                   *excelize.File
	odCopaRows, odIcbOrdRows [][]string
	odHeader                 map[string]int
	icbDbList                []icb_db_t
	dbList                   []db_t
}

func Exec(confFilePath string) error {
	data := &data_t{}

	log.Print("Reading... ")
	conf, err := readConfig(confFilePath)
	if err != nil {
		return err
	}
	data.conf = conf

	odXlsx, err := excelize.OpenFile(conf.files["OD"])
	if err != nil {
		return err
	}
	data.odXlsx = odXlsx

	odCopaRows := odXlsx.GetRows(conf.sheets["OD_COPA"])
	if len(odCopaRows) < 2 {
		return errors.New("Can not find 'COPA original data' sheet")
	}

	odIcbOrdRows := odXlsx.GetRows(conf.sheets["OD_ICB_ORD"])
	if len(odIcbOrdRows) < 2 {
		return errors.New("Can not find 'ICB_ORD' sheet")
	}

	odHeader, err := getODHeader(odCopaRows[0], odIcbOrdRows[0], conf)
	if err != nil {
		return err
	}
	data.odHeader = odHeader
	data.odCopaRows = odCopaRows[1:]
	data.odIcbOrdRows = odIcbOrdRows[1:]
	log.Println("OK!")

	log.Print("Calculating... ")
	splitIcbDb(data)
	resolveIcbDbRelation(data)
	handleUnusedDb(data)
	log.Println("OK!")

	log.Print("Outputing... ")
	output(data, "")
	log.Println("OK!")
	return nil
}

func splitIcbDb(data *data_t) {
	len := len(data.odCopaRows)
	icbDbList := make([]icb_db_t, 0, len/3*2)
	dbList := make([]db_t, 0, len/2)

	wbsColIdx := data.odHeader["OD_COPA_WBS"]
	soNoColIdx := data.odHeader["OD_COPA_SO"]
	tradPartnColIdx := data.odHeader["OD_COPA_TP"]
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
			icbDbList = append(icbDbList, icb_db_t{index, db_t{-1, 0}})
		} else {
			//DB
			dbList = append(dbList, db_t{index, 0})
		}
	}
	data.icbDbList, data.dbList = icbDbList, dbList
}

func resolveIcbDbRelation(data *data_t) {
	for icbIndex, icbDb := range data.icbDbList {
		if dbIndex, ok := findDbByODIcbOrdInDbList(data, icbDb.idx); ok {
			if dbIndex >= 0 {
				data.dbList[dbIndex].cnt++
				data.icbDbList[icbIndex].db = data.dbList[dbIndex]
			} else {
				data.icbDbList[icbIndex].db = db_t{-11, 0}
			}
		} else {
			//TODO shoud find in other table
		}
	}
}

//return: int: index of dblist; bool: whether it has found in OD_ICB_ORD
func findDbByODIcbOrdInDbList(data *data_t, rowIdx int) (int, bool) {
	soNoColIdx := data.odHeader["OD_COPA_SO"]
	icbOrdSoNoColIdx := data.odHeader["OD_ICB_ORD_SN"]
	icbOrdWbsColIdx := data.odHeader["OD_ICB_ORD_WBS"]
	icbOrdDbSoNoColIdx := data.odHeader["OD_ICB_ORD_DB_SN"]

	icbSoNo := strings.TrimSpace(data.odCopaRows[rowIdx][soNoColIdx])
	index := -1
	for _, odIcbOrdRow := range data.odIcbOrdRows {
		if icbSoNo == strings.TrimSpace(odIcbOrdRow[icbOrdSoNoColIdx]) {
			wbs := strings.TrimSpace(odIcbOrdRow[icbOrdWbsColIdx])
			if wbs != "" {
				index = matchODCopaWBS(data, wbs, rowIdx)
			} else {
				dbSoNo := strings.TrimSpace(odIcbOrdRow[icbOrdDbSoNoColIdx])
				index = matcODCopaSoNo(data, dbSoNo, rowIdx)
			}
			return index, true
		}
	}

	return index, false
}

func matchODCopaWBS(data *data_t, wbs string, rowIdx int) int {
	wbsColIdx := data.odHeader["OD_COPA_WBS"]
	productHierarchyColIdx := data.odHeader["OD_COPA_PH"]

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
	soNoColIdx := data.odHeader["OD_COPA_SO"]
	productHierarchyColIdx := data.odHeader["OD_COPA_PH"]

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

func handleUnusedDb(data *data_t) {
	for _, db := range data.dbList {
		if db.cnt == 0 {
			// have not a corresponding icb
			data.icbDbList = append(data.icbDbList, icb_db_t{-1, db})
		}
	}
}
