package copa

import (
	"errors"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type header_t struct {
	hierarchyIdx, businessIdx, exportIdx, tradPartnIdx, profitCenterIdx, partnerProfitCenterIdx int
}

func handleHeader(xlsx *excelize.File, conf *config_t) (header_t, error) {
	header := header_t{-1, -1, -1, -1, -1, -1}
	rows := xlsx.GetRows(conf.sheetName)
	if len(rows) == 0 {
		return header, errors.New("Can not find sheet:" + conf.sheetName)
	}

	for index, value := range rows[0] {
		title := strings.ToLower(value)
		switch title {
		case conf.columns["HIERARCHY"]:
			header.hierarchyIdx = index
		case conf.columns["BUSINESS"]:
			header.businessIdx = index
		case conf.columns["EXPORT"]:
			header.exportIdx = index
		case conf.columns["TRAD_PARTN"]:
			header.tradPartnIdx = index
		case conf.columns["PROFIT_CENTER"]:
			header.profitCenterIdx = index
		case conf.columns["PARTNER_PROFIT_CENTER"]:
			header.partnerProfitCenterIdx = index
		}
	}

	if header.tradPartnIdx < 0 {
		return header, errors.New("Can not find trad partn column")
	}

	if header.exportIdx < 0 {
		return header, errors.New("Can not find export column")
	}

	if header.profitCenterIdx < 0 {
		return header, errors.New("Can not find profit center column")
	}

	if header.partnerProfitCenterIdx < 0 {
		return header, errors.New("Can not find partner profit center column")
	}

	if header.businessIdx < 0 {
		xlsx.InsertCol(conf.sheetName, "A")
		xlsx.SetCellStr(conf.sheetName, "A1", "Business")

		header.businessIdx = 0

		header.exportIdx++
		header.tradPartnIdx++
		header.profitCenterIdx++
		header.partnerProfitCenterIdx++
	}

	if header.hierarchyIdx < 0 {
		xlsx.InsertCol(conf.sheetName, "A")
		xlsx.SetCellStr(conf.sheetName, "A1", "Hierarchy")

		header.hierarchyIdx = 0

		header.businessIdx++
		header.exportIdx++
		header.tradPartnIdx++
		header.profitCenterIdx++
		header.partnerProfitCenterIdx++
	}

	return header, nil
}
