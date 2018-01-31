package icbdb

import (
	"log"
	"strconv"
	"xlsx-processing/util"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const _OUTPUT_FILE = "res.xlsx"
const _OUTPUT_SHEET = "Sheet1"

type outputCols_t struct {
	PC, SO, TP, EX, NO, OOH, NS, COS, GM, DB_PC, DB_SO, DB_TP int
	DB_EX, DB_PH, DB_PPC, DB_NO, DB_OOH, DB_NS, DB_COS, DB_GM int
}

var outputCols outputCols_t = outputCols_t{0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19}

func output(data *data_t, path string) error {

	xlsx := excelize.NewFile()
	writeHeader(xlsx, data.conf)
	writeBody(xlsx, data)
	xlsx.SaveAs(path + _OUTPUT_FILE)
	return nil
}

func writeHeader(xlsx *excelize.File, conf *config_t) error {
	// OD_COPA_PC	Profit center
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.PC), conf.columns["OD_COPA_PC"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_PC), "[DB]"+conf.columns["OD_COPA_PC"])
	// OD_COPA_SO	Sales order
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.SO), conf.columns["OD_COPA_SO"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_SO), "[DB]"+conf.columns["OD_COPA_SO"])
	// OD_COPA_WBS	WBS Element
	// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, 2), conf.columns["OD_COPA_WBS"])
	// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, 15), "[DB]"+conf.columns["OD_COPA_WBS"])
	// OD_COPA_CUS	Customer
	// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, 3), conf.columns["OD_COPA_CUS"])
	// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, 16), "[DB]"+conf.columns["OD_COPA_CUS"])
	// OD_COPA_TP	Trad. partn.
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.TP), conf.columns["OD_COPA_TP"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_TP), "[DB]"+conf.columns["OD_COPA_TP"])
	// OD_COPA_EX	Export
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.EX), conf.columns["OD_COPA_EX"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_EX), "[DB]"+conf.columns["OD_COPA_EX"])
	// OD_COPA_PH	Product Hierarchy
	// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, 6), conf.columns["OD_COPA_PH"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_PH), "[DB]"+conf.columns["OD_COPA_PH"])
	// OD_COPA_PPC	Partner Profit Center:
	// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, 7), conf.columns["OD_COPA_PPC"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_PPC), "[DB]"+conf.columns["OD_COPA_PPC"])
	// OD_COPA_NO	New Order
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.NO), conf.columns["OD_COPA_NO"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_NO), "[DB]"+conf.columns["OD_COPA_NO"])
	// OD_COPA_OOH	Orders on hand
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.OOH), conf.columns["OD_COPA_OOH"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_OOH), "[DB]"+conf.columns["OD_COPA_OOH"])
	// OD_COPA_NS	Net Sales
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.NS), conf.columns["OD_COPA_NS"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_NS), "[DB]"+conf.columns["OD_COPA_NS"])
	// OD_COPA_COS	COS
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.COS), conf.columns["OD_COPA_COS"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_COS), "[DB]"+conf.columns["OD_COPA_COS"])
	// OD_COPA_GM	Gr. Margin
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.GM), conf.columns["OD_COPA_GM"])
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_GM), "[DB]"+conf.columns["OD_COPA_GM"])

	return nil
}

func writeBody(xlsx *excelize.File, data *data_t) error {
	warnCellStyle, _ := xlsx.NewStyle(`{"fill":{"type":"pattern","color":["#FFFF00"],"pattern":1}}`)
	rowIdx := 0
	for _, icbdb := range data.icbDbList {
		rowIdx++
		// OD_COPA_PC	Profit center
		ibcPCAxis := util.Axis(rowIdx, outputCols.PC)
		ibcPC := getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_PC")
		xlsx.SetCellStr(_OUTPUT_SHEET, ibcPCAxis, ibcPC)
		dbPCAxis := util.Axis(rowIdx, outputCols.DB_PC)
		dbPC := getDbValFromODCopa(data, icbdb.db, "OD_COPA_PC")
		xlsx.SetCellStr(_OUTPUT_SHEET, dbPCAxis, dbPC)
		// OD_COPA_SO	Sales order
		xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.SO), getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_SO"))

		if dbWBS := getDbValFromODCopa(data, icbdb.db, "OD_COPA_WBS"); dbWBS != "" {
			xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.SO), dbWBS[0:17])
		} else {
			xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.SO), getDbValFromODCopa(data, icbdb.db, "OD_COPA_SO"))
		}
		// OD_COPA_WBS	WBS Element
		// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, 2), getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_WBS"))
		// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, 15), getDbValFromODCopa(data, icbdb.db, "OD_COPA_WBS"))
		// OD_COPA_CUS	Customer
		// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, 3), getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_CUS"))
		// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, 16), getDbValFromODCopa(data, icbdb.db, "OD_COPA_CUS"))
		// OD_COPA_TP	Trad. partn.
		xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.TP), getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_TP"))
		xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_TP), getDbValFromODCopa(data, icbdb.db, "OD_COPA_TP"))
		// OD_COPA_EX	Export
		xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.EX), getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_EX"))
		xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_EX), getDbValFromODCopa(data, icbdb.db, "OD_COPA_EX"))
		// OD_COPA_PH	Product Hierarchy
		// ibcPHAxis := util.Axis(rowIdx, 6)
		// ibcPH := getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_PH")
		// xlsx.SetCellStr(_OUTPUT_SHEET, ibcPHAxis, ibcPH)
		dbPHAxis := util.Axis(rowIdx, outputCols.DB_PH)
		dbPH := getDbValFromODCopa(data, icbdb.db, "OD_COPA_PH")
		if dbPH == "" {
			dbPH = getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_PH")
		}
		xlsx.SetCellStr(_OUTPUT_SHEET, dbPHAxis, dbPH)
		// OD_COPA_PPC	Partner Profit Center:
		// xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, 7), getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_PPC"))
		xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_PPC), getDbValFromODCopa(data, icbdb.db, "OD_COPA_PPC"))
		// OD_COPA_NO	New Order
		xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.NO), getIcbFloatFromODCopa(data, icbdb.idx, "OD_COPA_NO"))
		xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_NO), getDbFloatFromODCopa(data, icbdb.db, "OD_COPA_NO"))
		// OD_COPA_OOH	Orders on hand
		xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.OOH), getIcbFloatFromODCopa(data, icbdb.idx, "OD_COPA_OOH"))
		xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_OOH), getDbFloatFromODCopa(data, icbdb.db, "OD_COPA_OOH"))
		// OD_COPA_NS	Net Sales
		xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.NS), getIcbFloatFromODCopa(data, icbdb.idx, "OD_COPA_NS"))
		xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_NS), getDbFloatFromODCopa(data, icbdb.db, "OD_COPA_NS"))
		// OD_COPA_COS	COS
		xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.COS), getIcbFloatFromODCopa(data, icbdb.idx, "OD_COPA_COS"))
		xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_COS), getDbFloatFromODCopa(data, icbdb.db, "OD_COPA_COS"))
		// OD_COPA_GM	Gr. Margin
		xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.GM), getIcbFloatFromODCopa(data, icbdb.idx, "OD_COPA_GM"))
		xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_GM), getDbFloatFromODCopa(data, icbdb.db, "OD_COPA_GM"))

		// if ok := checkPCPH(data, ibcPC, ibcPH); !ok {
		// 	xlsx.SetCellStyle(_OUTPUT_SHEET, ibcPCAxis, ibcPCAxis, warnCellStyle)
		// 	xlsx.SetCellStyle(_OUTPUT_SHEET, ibcPHAxis, ibcPHAxis, warnCellStyle)
		// }
		if ok := checkPCPH(data, dbPC, dbPH); !ok {
			xlsx.SetCellStyle(_OUTPUT_SHEET, dbPCAxis, dbPCAxis, warnCellStyle)
			xlsx.SetCellStyle(_OUTPUT_SHEET, dbPHAxis, dbPHAxis, warnCellStyle)
		}
	}

	numStyle, _ := xlsx.NewStyle(`{"custom_number_format": "#,##0.00_);[red](#,##0.00)"}`)
	xlsx.SetCellStyle(_OUTPUT_SHEET, util.Axis(1, 8), util.Axis(rowIdx, 12), numStyle)
	xlsx.SetCellStyle(_OUTPUT_SHEET, util.Axis(1, 21), util.Axis(rowIdx, 25), numStyle)
	return nil
}

func getIcbValFromODCopa(data *data_t, row int, colKey string) string {
	if row < 0 {
		return ""
	}

	return data.odCopaRows[row][data.odHeader[colKey]]
}

func getIcbFloatFromODCopa(data *data_t, row int, colKey string) float64 {
	return convertToFloat(getIcbValFromODCopa(data, row, colKey))
}

func getDbValFromODCopa(data *data_t, db db_t, colKey string) string {
	if db.idx < 0 {
		return ""
	}

	if db.cnt > 1 && (colKey == "OD_COPA_NO" ||
		colKey == "OD_COPA_OOH" ||
		colKey == "OD_COPA_NS" ||
		colKey == "OD_COPA_COS" ||
		colKey == "OD_COPA_GM") {
		return ""
	}

	return data.odCopaRows[db.idx][data.odHeader[colKey]]
}

func getDbFloatFromODCopa(data *data_t, db db_t, colKey string) float64 {
	return convertToFloat(getDbValFromODCopa(data, db, colKey))
}

func convertToFloat(val string) float64 {
	if val == "" {
		return 0
	}

	res, err := strconv.ParseFloat(val, 64)
	if err != nil {
		log.Println("Converting to float failed", val)
		log.Println(err)
	}
	return res
}

func checkPCPH(data *data_t, pc, ph string) bool {
	return true
}
