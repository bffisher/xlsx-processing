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

func output(data *data_t) error {

	xlsx := excelize.NewFile()
	writeHeader(xlsx, data.conf)
	writeBody(xlsx, data)
	xlsx.SaveAs(data.path + _OUTPUT_FILE)
	return nil
}

func writeHeader(xlsx *excelize.File, conf *config_t) error {
	// Profit center
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.PC), conf.columns["OD_COPA_PC"])
	// Sales order
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.SO), conf.columns["OD_COPA_SO"])
	// Trad. partn.
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.TP), conf.columns["OD_COPA_TP"])
	// Export
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.EX), conf.columns["OD_COPA_EX"])
	// New Order
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.NO), conf.columns["OD_COPA_NO"])
	// Orders on hand
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.OOH), conf.columns["OD_COPA_OOH"])
	// Net Sales
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.NS), conf.columns["OD_COPA_NS"])
	// COS
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.COS), conf.columns["OD_COPA_COS"])
	// Gr. Margin
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.GM), conf.columns["OD_COPA_GM"])

	//[DB]
	// Profit center
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_PC), "[DB]"+conf.columns["OD_COPA_PC"])
	// Sales order
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_SO), "[DB]"+conf.columns["OD_COPA_SO"])
	// Trad. partn.
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_TP), "[DB]"+conf.columns["OD_COPA_TP"])
	// Export
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_EX), "[DB]"+conf.columns["OD_COPA_EX"])
	// Product Hierarchy
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_PH), "[DB]"+conf.columns["OD_COPA_PH"])
	// Partner Profit Center:
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_PPC), "[DB]"+conf.columns["OD_COPA_PPC"])
	// New Order
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_NO), "[DB]"+conf.columns["OD_COPA_NO"])
	// Orders on hand
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.OOH), conf.columns["OD_COPA_OOH"])
	// Orders on hand
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_OOH), "[DB]"+conf.columns["OD_COPA_OOH"])
	// Net Sales
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_NS), "[DB]"+conf.columns["OD_COPA_NS"])
	// COS
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_COS), "[DB]"+conf.columns["OD_COPA_COS"])
	// Gr. Margin
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.DB_GM), "[DB]"+conf.columns["OD_COPA_GM"])

	return nil
}

func writeBody(xlsx *excelize.File, data *data_t) error {
	warnCellStyle, _ := xlsx.NewStyle(`{"fill":{"type":"pattern","color":["#FFFF00"],"pattern":1}}`)
	rowIdx := 0
	for _, icbdb := range data.icbDbList {
		rowIdx++
		if icbdb.idx >= 0 {
			// Profit center
			xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.PC), getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_PC"))
			// Sales order
			xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.SO), getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_SO"))
			// Trad. partn.
			xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.TP), getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_TP"))
			// Export
			xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.EX), getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_EX"))
			// New Order
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.NO), getIcbFloatFromODCopa(data, icbdb.idx, "OD_COPA_NO"))
			// Orders on hand
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.OOH), getIcbFloatFromODCopa(data, icbdb.idx, "OD_COPA_OOH"))
			// Net Sales
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.NS), getIcbFloatFromODCopa(data, icbdb.idx, "OD_COPA_NS"))
			// COS
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.COS), getIcbFloatFromODCopa(data, icbdb.idx, "OD_COPA_COS"))
			// Gr. Margin
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.GM), getIcbFloatFromODCopa(data, icbdb.idx, "OD_COPA_GM"))
		}

		if icbdb.db.idx >= 0 {
			// Profit center
			dbPCAxis := util.Axis(rowIdx, outputCols.DB_PC)
			dbPC := getDbValFromODCopa(data, icbdb.db, "OD_COPA_PC")
			xlsx.SetCellStr(_OUTPUT_SHEET, dbPCAxis, dbPC)
			// Sales order
			if dbWBS := getDbValFromODCopa(data, icbdb.db, "OD_COPA_WBS"); dbWBS != "" {
				xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_SO), dbWBS[0:17])
			} else {
				xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_SO), getDbValFromODCopa(data, icbdb.db, "OD_COPA_SO"))
			}
			// Trad. partn.
			xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_TP), getDbValFromODCopa(data, icbdb.db, "OD_COPA_TP"))
			// Export
			xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_EX), getDbValFromODCopa(data, icbdb.db, "OD_COPA_EX"))
			// 	Product Hierarchy
			dbPHAxis := util.Axis(rowIdx, outputCols.DB_PH)
			dbPH := getDbValFromODCopa(data, icbdb.db, "OD_COPA_PH")
			if dbPH == "" {
				dbPH = getIcbValFromODCopa(data, icbdb.idx, "OD_COPA_PH")
			}
			xlsx.SetCellStr(_OUTPUT_SHEET, dbPHAxis, dbPH)
			// Partner Profit Center:
			xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_PPC), getDbValFromODCopa(data, icbdb.db, "OD_COPA_PPC"))

			if icbdb.db.cnt <= 1 {
				// New Order
				xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_NO), getDbFloatFromODCopa(data, icbdb.db, "OD_COPA_NO"))
				// Orders on hand
				xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_OOH), getDbFloatFromODCopa(data, icbdb.db, "OD_COPA_OOH"))
				// Net Sales
				xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_NS), getDbFloatFromODCopa(data, icbdb.db, "OD_COPA_NS"))
				// COS
				xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_COS), getDbFloatFromODCopa(data, icbdb.db, "OD_COPA_COS"))
				// Gr. Margin
				xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_GM), getDbFloatFromODCopa(data, icbdb.db, "OD_COPA_GM"))
			}

			if ok := checkPCPH(data, dbPC, dbPH); !ok {
				xlsx.SetCellStyle(_OUTPUT_SHEET, dbPCAxis, dbPCAxis, warnCellStyle)
				xlsx.SetCellStyle(_OUTPUT_SHEET, dbPHAxis, dbPHAxis, warnCellStyle)
			}
		}
	}

	numStyle, _ := xlsx.NewStyle(`{"custom_number_format": "#,##0.00_);[red](#,##0.00)"}`)
	xlsx.SetCellStyle(_OUTPUT_SHEET, util.Axis(1, 8), util.Axis(rowIdx, 12), numStyle)
	xlsx.SetCellStyle(_OUTPUT_SHEET, util.Axis(1, 21), util.Axis(rowIdx, 25), numStyle)
	return nil
}

func getIcbValFromODCopa(data *data_t, row int, colKey string) string {
	return data.odCopaRows[row][data.odCopaHeader[colKey]]
}

func getIcbFloatFromODCopa(data *data_t, row int, colKey string) float64 {
	return convertToFloat(getIcbValFromODCopa(data, row, colKey))
}

func getDbValFromODCopa(data *data_t, db db_t, colKey string) string {
	return data.odCopaRows[db.idx][data.odCopaHeader[colKey]]
}

func getDbFloatFromODCopa(data *data_t, db db_t, colKey string) float64 {
	return convertToFloat(getDbValFromODCopa(data, db, colKey))
}

func convertToFloat(val string) float64 {
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
