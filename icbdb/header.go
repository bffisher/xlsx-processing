package icbdb

import (
	"strings"
	"xlsx-processing/util"
)

func getODHeader(copaRow []string, icbOrdRow []string, conf *config_t) (map[string]int, error) {
	header := make(map[string]int)
	util.ConvColNameToIdx(copaRow, conf.columns, header, func(key string) bool {
		return strings.Contains(key, "OD_COPA")
	})
	util.ConvColNameToIdx(icbOrdRow, conf.columns, header, func(key string) bool {
		return strings.Contains(key, "OD_ICB_ORD")
	})
	return header, util.CheckColNameIdx(conf.columns, header)
}
