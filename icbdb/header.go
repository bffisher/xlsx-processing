package icbdb

import (
	"strings"
	"xlsx-processing/util"
)

func getODCopaHeader(row []string, conf *config_t) (map[string]int, error) {
	header := make(map[string]int)
	filter := func(key string) bool {
		return strings.Contains(key, "OD_COPA")
	}
	util.ConvColNameToIdx(row, conf.columns, header, filter)
	return header, util.CheckColNameIdx(conf.columns, header, filter)
}

func getODIcbOrdHeader(row []string, conf *config_t) (map[string]int, error) {
	header := make(map[string]int)
	filter := func(key string) bool {
		return strings.Contains(key, "OD_ICB_ORD")
	}
	util.ConvColNameToIdx(row, conf.columns, header, filter)
	return header, util.CheckColNameIdx(conf.columns, header, filter)
}
