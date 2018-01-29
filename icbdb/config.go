package icbdb

import (
	"xlsx-processing/util"
)

type config_t struct {
	files, sheets, columns, products map[string]string
}

func handleConfig(filePath string) (config_t, error) {
	conf := config_t{make(map[string]string), make(map[string]string), make(map[string]string), make(map[string]string)}

	rows, err := util.ReadConfig(filePath)
	if err != nil {
		return conf, err
	}

	for _, row := range rows {
		util.CopyRowToMap(&conf.files, row, 0, 1)

		util.CopyRowToMap(&conf.sheets, row, 3, 4)

		util.CopyRowToMap(&conf.columns, row, 6, 7)

		util.CopyRowToMapWithFunc(&conf.columns, row, 9, 10, func(key, value string) (string, string) {
			_, key = util.SplitCodeName(key)
			return key, value
		})
	}

	return conf, nil
}
