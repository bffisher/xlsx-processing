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
	DB_PC, DB_SO, DB_TP, DB_EX, DB_PH, DB_PPC, DB_NO, DB_OOH, DB_NS, DB_COS, DB_GM int
	PC, SO, TP, EX, NO, OOH, NS, COS, GM                                           int
	NO_ICB ,PRODUCT, PROVINCE, CLASSFICATION int
}

var outputCols outputCols_t = outputCols_t{0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
	11, 12, 13, 14, 15, 16, 17, 18, 19, 20,	21, 22, 23}
var lastOutputCol int = outputCols.CLASSFICATION

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

	// PRODUCT
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.PRODUCT), "PRODUCT")
	// PROVINCE
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.PROVINCE), "PROVINCE")
	// CLASSFICATION
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(0, outputCols.CLASSFICATION), "CLASSFICATION")


	headerStyle, _ := xlsx.NewStyle(`{"fill":{"type":"pattern","color":["#000080"],"pattern":1}, "font":{"color":"#FFFFFF", "bold":true}}`)
	xlsx.SetCellStyle(_OUTPUT_SHEET, util.Axis(0, 0), util.Axis(0, lastOutputCol), headerStyle)
	return nil
}

func writeBody(xlsx *excelize.File, data *data_t) error {
	warnCellStyle, _ := xlsx.NewStyle(`{"fill":{"type":"pattern","color":["#FFFF00"],"pattern":1}}`)
	rowIdx := 1
	for _, db := range data.dbList {
		writeDb(xlsx, data, rowIdx, &db, warnCellStyle)

		for _, idx := range db.icbIdxs {
			writeIcb(xlsx, data, rowIdx, idx)
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.PROVINCE), db.province)
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.CLASSFICATION), db.classfication)
			rowIdx++
		}

		if len(db.icbIdxs) == 0 {
			//Product
			writeProduct(xlsx, data, rowIdx, db.idx)
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.PROVINCE), db.province)
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.CLASSFICATION), db.classfication)
			rowIdx++
		}
	}

	for _, icb := range data.icbList {
		if icb.dbIdx == _DB_IDX_NO_DATA || icb.dbIdx == _DB_IDX_NO_RELATION{
			writeIcb(xlsx, data, rowIdx, icb.idx)
			if icb.dbIdx == _DB_IDX_NO_DATA {
				// Sales order
				if icb.dbWbs != ""{
					xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_SO), icb.dbWbs[0:17])
				}else{
					xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_SO), icb.dbSoNo)
				}
				// New Order
				xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_NO), 0)
				// Orders on hand
				xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_OOH), 0)
				// Net Sales
				xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_NS), 0)
				// COS
				xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_COS), 0)
				// Gr. Margin
				xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_GM), 0)
			}
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.PROVINCE), icb.province)
			xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.CLASSFICATION), icb.classfication)
			rowIdx++
		}
	}

	numStyle, _ := xlsx.NewStyle(`{"custom_number_format": "#,##0_);[red](#,##0)"}`)
	xlsx.SetCellStyle(_OUTPUT_SHEET, util.Axis(1, outputCols.NO), util.Axis(rowIdx, outputCols.GM), numStyle)
	xlsx.SetCellStyle(_OUTPUT_SHEET, util.Axis(1, outputCols.DB_NO), util.Axis(rowIdx, outputCols.DB_GM), numStyle)

	return nil
}

func writeDb(xlsx *excelize.File, data *data_t, rowIdx int, db *db_t, warnCellStyle int) {
	icbLen := len(db.icbIdxs)
	// Profit center
	dbPCAxis := util.Axis(rowIdx, outputCols.DB_PC)
	dbPC := getValFromODCopa(data, db.idx, "OD_COPA_PC")
	xlsx.SetCellStr(_OUTPUT_SHEET, dbPCAxis, dbPC)
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_PC)
	// Sales order
	if dbWBS := getValFromODCopa(data, db.idx, "OD_COPA_WBS"); dbWBS != "" {
		xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_SO), dbWBS[0:17])
	} else {
		xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_SO), getValFromODCopa(data, db.idx, "OD_COPA_SO"))
	}
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_SO)
	// Trad. partn.
	dbTR := getValFromODCopa(data, db.idx, "OD_COPA_TP")
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_TP), dbTR)
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_TP)
	// Export
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_EX), getValFromODCopa(data, db.idx, "OD_COPA_EX"))
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_EX)
	// 	Product Hierarchy
	dbPHAxis := util.Axis(rowIdx, outputCols.DB_PH)
	dbPH := getValFromODCopa(data, db.idx, "OD_COPA_PH")
	if dbPH == "" && len(db.icbIdxs) > 0 {
		dbPH = getValFromODCopa(data, db.icbIdxs[0], "OD_COPA_PH")
	}
	xlsx.SetCellStr(_OUTPUT_SHEET, dbPHAxis, dbPH)
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_PH)
	// Partner Profit Center:
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_PPC), getValFromODCopa(data, db.idx, "OD_COPA_PPC"))
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_PPC)
	// New Order
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_NO), getFloatFromODCopa(data, db.idx, "OD_COPA_NO"))
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_NO)
	// Orders on hand
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_OOH), getFloatFromODCopa(data, db.idx, "OD_COPA_OOH"))
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_OOH)
	// Net Sales
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_NS), getFloatFromODCopa(data, db.idx, "OD_COPA_NS"))
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_NS)
	// COS
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_COS), getFloatFromODCopa(data, db.idx, "OD_COPA_COS"))
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_COS)
	// Gr. Margin
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.DB_GM), getFloatFromODCopa(data, db.idx, "OD_COPA_GM"))
	tryMergeCells(xlsx, rowIdx, icbLen, outputCols.DB_GM)

	if ok := checkPCPH(data, dbPC, dbPH); !ok {
		xlsx.SetCellStyle(_OUTPUT_SHEET, dbPCAxis, dbPCAxis, warnCellStyle)
		xlsx.SetCellStyle(_OUTPUT_SHEET, dbPHAxis, dbPHAxis, warnCellStyle)
	}

	if icbLen == 0{
		if pc,_ := util.SplitCodeName(dbPC); pc == _PC_P8251 || pc == _PC_P8211{
			if tr,_ := util.SplitCodeName(dbTR); tr != _TR_4611{
				xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.NO_ICB), "NO ICB")
			}
		} 
	}
}

func writeIcb(xlsx *excelize.File, data *data_t, rowIdx, idx int) {
	// Profit center
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.PC), getValFromODCopa(data, idx, "OD_COPA_PC"))
	// Sales order
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.SO), getValFromODCopa(data, idx, "OD_COPA_SO"))
	// Trad. partn.
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.TP), getValFromODCopa(data, idx, "OD_COPA_TP"))
	// Export
	xlsx.SetCellStr(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.EX), getValFromODCopa(data, idx, "OD_COPA_EX"))
	// New Order
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.NO), getFloatFromODCopa(data, idx, "OD_COPA_NO"))
	// Orders on hand
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.OOH), getFloatFromODCopa(data, idx, "OD_COPA_OOH"))
	// Net Sales
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.NS), getFloatFromODCopa(data, idx, "OD_COPA_NS"))
	// COS
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.COS), getFloatFromODCopa(data, idx, "OD_COPA_COS"))
	// Gr. Margin
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.GM), getFloatFromODCopa(data, idx, "OD_COPA_GM"))
	//Product
	writeProduct(xlsx, data, rowIdx, idx)	
}

func writeProduct(xlsx *excelize.File, data *data_t, rowIdx, idx int){
	//Product
	_,productHierarchy := util.SplitCodeName(getValFromODCopa(data, idx, "OD_COPA_PH"))
	xlsx.SetCellValue(_OUTPUT_SHEET, util.Axis(rowIdx, outputCols.PRODUCT), productHierarchy)
}

func getValFromODCopa(data *data_t, row int, colKey string) string {
	return data.odCopaRows[row][data.odCopaHeader[colKey]]
}

func getFloatFromODCopa(data *data_t, row int, colKey string) float64 {
	return convertToFloat(getValFromODCopa(data, row, colKey))
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
	//TODO:
	return true
}

func tryMergeCells(xlsx *excelize.File, startRow, len, col int) {
	if len <= 1 {
		return
	}

	hcell := util.Axis(startRow, col)
	vcell := util.Axis(startRow+len-1, col)
	xlsx.MergeCell(_OUTPUT_SHEET, hcell, vcell)
}
