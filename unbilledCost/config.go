package unbilledCost

import (
	"xlsx-processing/util"
)

type config_t struct {
	files, sheets, columns map[string]string
	gis2Sheets            [][2]string
}

func readConfig(filePath string) (*config_t, error) {
	conf := config_t{make(map[string]string), make(map[string]string), make(map[string]string), make([][2]string, 0, 10)}

	rows, err := util.ReadConfig(filePath)
	if err != nil {
		return &conf, err
	}

	for _, row := range rows {
		util.CopyRowToMap(conf.files, row, 0, 1)
		util.CopyRowToMap(conf.sheets, row, 3, 4)
		util.CopyRowToMap(conf.columns, row, 6, 7)
		conf.gis2Sheets = util.CopyRowToArray(conf.gis2Sheets, row, 9, 10)
	}

	return &conf, nil
}
