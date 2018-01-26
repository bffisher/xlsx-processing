package copa

import (
	"strings"
	"xlsx-processing/util"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func handleBody(xlsx *excelize.File, conf *config_t, header *header_t) {
	rows := xlsx.GetRows(conf.sheetName)

	for index, row := range rows[1:] {
		//add header row
		index++

		export := strings.TrimSpace(row[header.exportIdx])
		trad := extractCode(row[header.exportIdx])
		profitCenter := extractCode(row[header.profitCenterIdx])
		partnerProfitCenter := extractCode(row[header.partnerProfitCenterIdx])

		xlsx.SetCellStr(conf.sheetName, util.Axis(index, header.hierarchyIdx), conf.hierarchyValues[profitCenter])

		businessKey := calculateBusiness(export, trad, profitCenter, partnerProfitCenter)
		if businessKey != "" {
			xlsx.SetCellStr(conf.sheetName, util.Axis(index, header.businessIdx), conf.businessValues[businessKey])
		}
	}
}

func extractCode(value string) string {
	return strings.TrimSpace(strings.Split(value, ":")[0])
}

func calculateBusiness(export, trad, profitCenter, partnerProfitCenter string) string {
	if export == "Y" { //[Export] == Y ?
		return "EXPORT_Y"
	} else if trad == "004611" { //[Trad. partn.] == 004611
		if profitCenter == "P8251" { //[Profit center] == P8251
			return "TP_004611_PC_P8251"
		} else {
			return "TP_004611_PC_NOT_P8251"
		}
	} else if partnerProfitCenter == "0000000000" { //[Partner Profit Center.] == 0000000000
		return "PPC_0000000000"
	} else if partnerProfitCenter == "SEM" { //[Partner Profit Center.] == SEM
		if trad == "005531" { //[Trad. partn.]  == 005531
			return "PPC_SEM_TP_005531"
		} else {
			return "PPC_SEM_TP_NOT_005531"
		}
	} else if export != "" {
		return "OTHER"
	} else {
		return ""
	}
}
