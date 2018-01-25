package copa

import (
	"errors"
	"strings"
	"xlsx-processing/util"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const COL_NAME_HIERARCHY = "hierarchy"
const COL_NAME_BUSINESS = "business"
const COL_NAME_EXPORT = "export"
const COL_NAME_TRAD_PARTN = "trad. partn."
const COL_NAME_PROFIT_CENTER = "profit center"
const COL_NAME_PARTNER_PROFIT_CENTER = "partner profit center"

var profitCenterHierarchy map[string]string = map[string]string{
	// Hierarchy	Profit center
	// LP	        P82LPPO
	"P82LPPO": "LP",
	// MS CS	    P82MSTS
	"P82MSTS": "MS CS",
	// MS S_LD	  P82MSSLD
	"P82MSSLD": "MS S_LD",
	// MS S_CB	  P82MSSCB
	"P82MSSCB": "MS S_CB",
	// MS S_PSS	  P82491
	"P82491": "MS S_PSS",
	// MS O_VI	  P8251
	"P8251": "MS O_VI",
	// MS O_LD	  P8211
	"P8211": "MS O_LD",
	// MS O_CB	  P8221
	"P8221": "MS O_CB",
}

type header_t struct {
	hierarchyIdx, businessIdx, exportIdx, tradPartnIdx, profitCenterIdx, partnerProfitCenterIdx int
}

func Exec(filePath, sheetName string) error {
	xlsx, err := excelize.OpenFile(filePath)
	if err != nil {
		return err
	}

	header, err := handleHeader(xlsx, sheetName)
	if err != nil {
		return err
	}

	handleBody(xlsx, sheetName, &header)

	return xlsx.SaveAs(strings.Replace(filePath, ".xlsx", "_res.xlsx", -1))

}

func handleHeader(xlsx *excelize.File, sheetName string) (header_t, error) {
	header := header_t{-1, -1, -1, -1, -1, -1}
	rows := xlsx.GetRows(sheetName)
	if len(rows) == 0 {
		return header, errors.New("Can not find sheet:" + sheetName)
	}

	for index, value := range rows[0] {
		title := strings.ToLower(value)
		switch title {
		case COL_NAME_HIERARCHY:
			header.hierarchyIdx = index
		case COL_NAME_BUSINESS:
			header.businessIdx = index
		case COL_NAME_EXPORT:
			header.exportIdx = index
		case COL_NAME_TRAD_PARTN:
			header.tradPartnIdx = index
		case COL_NAME_PROFIT_CENTER:
			header.profitCenterIdx = index
		case COL_NAME_PARTNER_PROFIT_CENTER:
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
		xlsx.InsertCol(sheetName, "A")
		xlsx.SetCellStr(sheetName, "A1", "Business")

		header.businessIdx = 0

		header.exportIdx++
		header.tradPartnIdx++
		header.profitCenterIdx++
		header.partnerProfitCenterIdx++
	}

	if header.hierarchyIdx < 0 {
		xlsx.InsertCol(sheetName, "A")
		xlsx.SetCellStr(sheetName, "A1", "Hierarchy")

		header.hierarchyIdx = 0

		header.businessIdx++
		header.exportIdx++
		header.tradPartnIdx++
		header.profitCenterIdx++
		header.partnerProfitCenterIdx++
	}

	return header, nil
}

func handleBody(xlsx *excelize.File, sheetName string, header *header_t) {
	rows := xlsx.GetRows(sheetName)

	for index, row := range rows[1:] {
		//Skip header row
		index++

		export := strings.TrimSpace(row[header.exportIdx])
		trad := extractCode(row[header.exportIdx])
		profitCenter := extractCode(row[header.profitCenterIdx])
		partnerProfitCenter := extractCode(row[header.partnerProfitCenterIdx])

		xlsx.SetCellStr(sheetName, util.Axis(index, header.hierarchyIdx), profitCenterHierarchy[profitCenter])

		business := calculateBusiness(export, trad, profitCenter, partnerProfitCenter)
		xlsx.SetCellStr(sheetName, util.Axis(index, header.businessIdx), business)
	}
}

func extractCode(value string) string {
	return strings.TrimSpace(strings.Split(value, ":")[0])
}

func calculateBusiness(export, trad, profitCenter, partnerProfitCenter string) string {
	if export == "Y" { //[Export] == Y ?
		//["B"] = Export
		return "Export"
	} else if trad == "004611" { //[Trad. partn.] == 004611
		if profitCenter == "P8251" { //[Profit center] == P8251
			//["B"] = domestic MS ICB other BU
			return "domestic MS ICB other BU"
		} else {
			//["B"] = domestic MS ICB own BU
			return "domestic MS ICB own BU"
		}
	} else if partnerProfitCenter == "0000000000" { //[Partner Profit Center.] == 0000000000
		//["B"] = Domestic 3rd party
		return "Domestic 3rd party"
	} else if partnerProfitCenter == "SEM" { //[Partner Profit Center.] == SEM
		if trad == "005531" { //[Trad. partn.]  == 005531
			//["B"] = domestic own division own BU
			return "domestic own division own BU"
		} else {
			//["B"] = domestic own division other BU
			return "domestic own division other BU"
		}
	} else {
		//["B"] = Domestic Inter-company (other division)
		return "Domestic Inter-company (other division)"
	}
}
