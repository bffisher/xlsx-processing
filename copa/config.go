package copa

import (
	"errors"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type config_t struct {
	filePath, sheetName                      string
	columns, hierarchyValues, businessValues map[string]string
}

func handleConfig(filePath string) (config_t, error) {
	conf := config_t{"", "", make(map[string]string), make(map[string]string), make(map[string]string)}

	xlsx, err := excelize.OpenFile(filePath)
	if err != nil {
		return conf, err
	}

	rows := xlsx.GetRows(xlsx.GetSheetName(1))
	if len(rows) < 2 || len(rows[0]) < 11 {
		return conf, errors.New("Read config info failed.")
	}

	for index, row := range rows[1:] {
		if index == 0 {
			conf.filePath = row[0]
			conf.sheetName = row[1]
		}

		colId, colName := row[3], row[4]
		if colId != "" {
			conf.columns[colId] = strings.ToLower(colName)
		}

		profitCenter, hierarchy := row[6], row[7]
		if profitCenter != "" {
			conf.hierarchyValues[profitCenter] = hierarchy
		}

		businessId, businessValue := row[9], row[10]
		if businessId != "" {
			conf.businessValues[businessId] = businessValue
		}
	}

	return conf, nil
}
