package icbdb

import (
	"errors"
	"strings"
	"xlsx-processing/util"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type data_t struct {
	conf                     *config_t
	odCopaRows, odIcbOrdRows [][]string
	odCopaHeader             *odCopaHeader_t
	odIcbOrdHeader           *odIcbOrdHeader_t
	icbList, dbList          []int
	icbDbRelation            map[int]int
}

func Exec(confFilePath string) error {
	data := &data_t{}

	conf, err := handleConfig(confFilePath)
	if err != nil {
		return err
	}
	data.conf = &conf

	odXlsx, err := excelize.OpenFile(conf.files["OD_FILE"])
	if err != nil {
		return err
	}

	odCopaRows := odXlsx.GetRows(conf.sheets["OD_COPA"])
	if len(odCopaRows) < 2 {
		return errors.New("Can not find 'COPA original data' sheet")
	}
	odCopaHeader, err := handleODCopaHeader(odCopaRows[0], &conf)
	if err != nil {
		return err
	}
	data.odCopaRows = odCopaRows[1:]
	data.odCopaHeader = &odCopaHeader

	odIcbOrdRows := odXlsx.GetRows(conf.sheets["OD_ICB_ORD"])
	if len(odIcbOrdRows) < 2 {
		return errors.New("Can not find 'ICB_ORD' sheet")
	}
	odIcbOrdHeader, err := handleODIcbOrdHeader(odIcbOrdRows[0], &conf)
	if err != nil {
		return err
	}
	data.odIcbOrdRows = odIcbOrdRows[1:]
	data.odIcbOrdHeader = &odIcbOrdHeader

	splitIcbDb(data)
	resolveIcbDbRelation(data)

	return nil
}

func splitIcbDb(data *data_t) {
	cap := len(data.odCopaRows) / 2
	icbList := make([]int, 0, cap)
	dbList := make([]int, 0, cap)

	for index, row := range data.odCopaRows {
		if wbs := row[data.odCopaHeader.wbsIdx]; wbs != "" {
			_, wbs = util.SplitCodeName(wbs)
			if strings.Contains(strings.ToUpper(wbs), "POC") {
				continue
			}
		}

		soNo := row[data.odCopaHeader.soNoIdx]
		tradPartn, _ := util.SplitCodeName(row[data.odCopaHeader.tradPartnIdx])
		if soNo != "" && tradPartn == "004611" {
			//ICB
			icbList = append(icbList, index)
		} else {
			//DB
			dbList = append(dbList, index)
		}
	}
	data.icbList, data.dbList = icbList, dbList
}

func resolveIcbDbRelation(data *data_t) {
	relation := make(map[int]int)
	for _, icbIdx := range data.icbList {

		if dbIdx, ok := findDbIdxByODIcbOrd(data, icbIdx); ok {
			relation[icbIdx] = dbIdx
		} else {
			//TODO shoud find in other table
			relation[icbIdx] = -99
		}
	}

	data.icbDbRelation = relation
}

func findDbIdxByODIcbOrd(data *data_t, icbIdx int) (int, bool) {
	icbSoNo := strings.TrimSpace(data.odCopaRows[icbIdx][data.odCopaHeader.soNoIdx])
	for _, odIcbOrdRow := range data.odIcbOrdRows {
		if icbSoNo == strings.TrimSpace(odIcbOrdRow[data.odIcbOrdHeader.icbSoNoIdx]) {
			wbs := strings.TrimSpace(odIcbOrdRow[data.odIcbOrdHeader.wbsIdx])
			dbIdx := -1
			if wbs != "" {
				dbIdx = matchODCopaWBS(data, wbs, icbIdx)
			} else {
				dbSoNo := strings.TrimSpace(odIcbOrdRow[data.odIcbOrdHeader.dbSoNoIdx])
				dbIdx = matcODCopaSoNo(data, dbSoNo, icbIdx)
			}

			return dbIdx, true
		}
	}

	return -19, false
}

func matchODCopaWBS(data *data_t, wbs string, icbIdx int) int {
	for _, dbIdx := range data.dbList {
		dbWbs := strings.TrimSpace(data.odCopaRows[dbIdx][data.odCopaHeader.wbsIdx])
		if dbWbs != "" && wbs[0:17] == dbWbs[0:17] {
			icbProductHierarchy := strings.TrimSpace(data.odCopaRows[icbIdx][data.odCopaHeader.productHierarchyIdx])
			dbProductHierarchy := strings.TrimSpace(data.odCopaRows[dbIdx][data.odCopaHeader.productHierarchyIdx])
			if matchProductHierarchy(data, icbProductHierarchy, dbProductHierarchy) {
				return dbIdx
			}
		}
	}
	return -11
}

func matcODCopaSoNo(data *data_t, dbSoNo string, icbIdx int) int {
	for _, dbIdx := range data.dbList {
		if dbSoNo == strings.TrimSpace(data.odCopaRows[dbIdx][data.odCopaHeader.soNoIdx]) {
			icbProductHierarchy := strings.TrimSpace(data.odCopaRows[icbIdx][data.odCopaHeader.productHierarchyIdx])
			dbProductHierarchy := strings.TrimSpace(data.odCopaRows[dbIdx][data.odCopaHeader.productHierarchyIdx])
			if matchProductHierarchy(data, icbProductHierarchy, dbProductHierarchy) {
				return dbIdx
			}
		}
	}
	return -12
}

func matchProductHierarchy(data *data_t, val1, val2 string) bool {
	_, name1 := util.SplitCodeName(val1)
	_, name2 := util.SplitCodeName(val2)

	return data.conf.products[name1] == data.conf.products[name2]
}
