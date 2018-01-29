package icbdb

import (
	"errors"
	"strings"
)

type odCopaHeader_t struct {
	wbsIdx, tradPartnIdx, soNoIdx, productHierarchyIdx int
}

func handleODCopaHeader(row []string, conf *config_t) (odCopaHeader_t, error) {
	header := odCopaHeader_t{-1, -1, -1, -1}
	for index, value := range row {
		title := strings.TrimSpace(value)
		switch title {
		case conf.columns["OD_COPA_WBS"]:
			header.wbsIdx = index
		case conf.columns["OD_COPA_TP"]:
			header.tradPartnIdx = index
		case conf.columns["OD_COPA_SONO"]:
			header.soNoIdx = index
		case conf.columns["OD_COPA_PH"]:
			header.productHierarchyIdx = index
		}
	}

	if header.tradPartnIdx < 0 {
		return header, errors.New("Can not find 'trad partn' column")
	}

	if header.wbsIdx < 0 {
		return header, errors.New("Can not find 'WBS Element' column")
	}

	if header.soNoIdx < 0 {
		return header, errors.New("Can not find 'Sales order' column")
	}

	if header.productHierarchyIdx < 0 {
		return header, errors.New("Can not find 'Product Hierarchy' column")
	}

	return header, nil
}

type odIcbOrdHeader_t struct {
	icbSoNoIdx, dbSoNoIdx, wbsIdx int
}

func handleODIcbOrdHeader(row []string, conf *config_t) (odIcbOrdHeader_t, error) {
	header := odIcbOrdHeader_t{-1, -1, -1}

	for index, value := range row {
		title := strings.TrimSpace(value)
		switch title {
		case conf.columns["OD_ICB_ORD_SONO"]:
			if header.icbSoNoIdx == -1 {
				//first So No column
				header.icbSoNoIdx = index
			} else if header.dbSoNoIdx == -1 {
				//second So No column
				header.dbSoNoIdx = index
			}
		case conf.columns["OD_ICB_ORD_WBS"]:
			header.wbsIdx = index
		}
	}

	if header.icbSoNoIdx < 0 {
		return header, errors.New("Can not find 'So No' column(ICB)")
	}

	if header.dbSoNoIdx < 0 {
		return header, errors.New("Can not find 'So No' column(DB)")
	}

	if header.wbsIdx < 0 {
		return header, errors.New("Can not find 'WBS' column")
	}

	return header, nil
}
