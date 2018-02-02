package icbdb

import (
	"xlsx-processing/util"
)

type config_t struct {
	files, sheets, columns, products map[string]string
	gisSheets, mcSheets              [][2]string
}

func readConfig(filePath string) (*config_t, error) {
	conf := config_t{make(map[string]string), make(map[string]string), make(map[string]string),
		make(map[string]string), make([][2]string, 0, 10), make([][2]string, 0, 10)}

	rows, err := util.ReadConfig(filePath)
	if err != nil {
		return &conf, err
	}

	for _, row := range rows {
		util.CopyRowToMap(conf.files, row, 0, 1)

		util.CopyRowToMap(conf.sheets, row, 3, 4)

		util.CopyRowToMap(conf.columns, row, 6, 7)

		util.CopyRowToMapWithFunc(conf.products, row, 9, 10, func(key, value string) (string, string) {
			_, key = util.SplitCodeName(key)
			return key, value
		})

		conf.gisSheets = util.CopyRowToArray(conf.gisSheets, row, 12, 13)
		conf.mcSheets = util.CopyRowToArray(conf.mcSheets, row, 15, 16)
	}

	return &conf, nil
}
