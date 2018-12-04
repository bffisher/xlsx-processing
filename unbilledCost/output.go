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
	leftBeginColIdx := 0
	rightBeginColIdx := leftBeginColIdx + len(data.ucData[0])
	extColIdx := rightBeginColIdx +  len(data.ucData[0])

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
		writeExDataForLeft(xlsx, ucLeft, rowIdx, extColIdx)

		if len(ucLeft.rightIdxs) == 0 && len(ucLeft.noRightSoNoList) == 0{
			rowIdx ++
			continue
		}
		
		for _, ucRightIdx := range(ucLeft.rightIdxs){
			writeUcData(xlsx, ucRightIdx, rowIdx, rightBeginColIdx)
			rowIdx ++
		}

		for _, soNo := range(ucLeft.noRightSoNoList){
			xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, rightBeginColIdx + data.ucObjColIdx), soNo)
			rowIdx ++
		} 
	}

	orderUcRightData()
	for _, ucRight := range(data.ucRightData){
		if ucRight.leftIdx == DB_IDX_NO_RELATION || ucRight.leftIdx == DB_IDX_NO_DATA{
			xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, leftBeginColIdx + data.ucObjColIdx), ucRight.dbSoNo)		
			writeExDataForRight(xlsx, ucRight, rowIdx, extColIdx)
			writeUcData(xlsx, ucRight.idx, rowIdx, rightBeginColIdx)
			rowIdx ++
		}
	}
}

func orderUcRightData(){
	len := len(data.ucRightData)
	for i:= 0; i<len; i++{
		if data.ucRightData[i].leftIdx == DB_IDX_NO_RELATION || data.ucRightData[i].leftIdx == DB_IDX_NO_DATA{
			for j:= i + 1; j < len; j++{
				if data.ucRightData[j].leftIdx == DB_IDX_NO_RELATION || data.ucRightData[j].leftIdx == DB_IDX_NO_DATA{
					if data.ucRightData[i].dbSoNo > data.ucRightData[j].dbSoNo{
						tmp:=data.ucRightData[i]
						data.ucRightData[i] = data.ucRightData[j]
						data.ucRightData[j] = tmp
					}
				}
			}
		}
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

func writeExDataForLeft(xlsx *excelize.File, ucLeft uc_left_data_t, rowIdx , beginColIdx int){
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 0 + beginColIdx), ucLeft.product)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 1 + beginColIdx), ucLeft.projectName)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 2 + beginColIdx), ucLeft.customerNo)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 3 + beginColIdx), ucLeft.customerName)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 4 + beginColIdx), ucLeft.contractNo)
}

func writeExDataForRight(xlsx *excelize.File, ucRight uc_right_data_t, rowIdx , beginColIdx int){
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 0 + beginColIdx), ucRight.product)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 1 + beginColIdx), ucRight.projectName)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 2 + beginColIdx), ucRight.customerNo)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 3 + beginColIdx), ucRight.customerName)
	xlsx.SetCellStr(OUTPUT_SHEET, util.Axis(rowIdx, 4 + beginColIdx), ucRight.contractNo)
}