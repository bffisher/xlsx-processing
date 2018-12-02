package unbilledCost

import (
	"log"
	// "strconv"
	"xlsx-processing/util"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const OUTPUT_FILE = "res.xlsx"
const OUTPUT_SHEET = "Sheet1"

func output() {
	log.Println("Output...")
	xlsx := excelize.NewFile()
	writeBody(xlsx)
	xlsx.SaveAs(util.Env().FilePath + OUTPUT_FILE)
	log.Println("Output...OK!")
}

func writeBody(xlsx *excelize.File){
	extColIdx := 1
	leftBeginColIdx := 7
	rightBeginColIdx := leftBeginColIdx + len(data.ucData[0])

	//write header
	writeUcHeader(xlsx, leftBeginColIdx)
	writeUcHeader(xlsx, rightBeginColIdx)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(0, 0 + extColIdx), "product")
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(0, 1 + extColIdx), "projectName")
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(0, 2 + extColIdx), "customerNo")
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(0, 3 + extColIdx), "customerName")
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(0, 4 + extColIdx), "contractNo")

	rowIdx := 1

	for _, ucLeft := range(data.ucLeftData){
		writeUcData(xlsx, ucLeft.idx, rowIdx, leftBeginColIdx)
		writeExData(xlsx, ucLeft, rowIdx, extColIdx)

		if len(ucLeft.rightIdxs) == 0{
			rowIdx ++
			continue
		}
		
		for _, ucRightIdx := range(ucLeft.rightIdxs){
			writeUcData(xlsx, ucRightIdx, rowIdx, rightBeginColIdx)
			rowIdx ++
		}
	}

	for _, ucRight := range(data.ucRightData){
		if ucRight.leftIdx != DB_IDX_NO_RELATION{
			continue
		}

		writeUcData(xlsx, ucRight.idx, rowIdx, rightBeginColIdx)
		rowIdx ++
	}
}

func writeUcHeader(xlsx *excelize.File, beginColIdx int){
	offColIdx := 0
	for _, cell := range(data.ucHeader){
		xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(0, offColIdx + beginColIdx), cell)

		offColIdx ++
	}
}

func writeUcData(xlsx *excelize.File, dataIdx, rowIdx , beginColIdx int){
	offColIdx := 0
	for _, cell := range(data.ucData[dataIdx]){
		xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, offColIdx + beginColIdx), cell)

		offColIdx ++
	}
}

func writeExData(xlsx *excelize.File, ucLeft uc_left_data_t, rowIdx , beginColIdx int){
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 0 + beginColIdx), ucLeft.product)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 1 + beginColIdx), ucLeft.projectName)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 2 + beginColIdx), ucLeft.customerNo)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 3 + beginColIdx), ucLeft.customerName)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 4 + beginColIdx), ucLeft.contractNo)
}