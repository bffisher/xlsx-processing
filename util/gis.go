package util

import(
	"errors"
	"strings"
	"strconv"
	"log"
	"github.com/360EntSecGroup-Skylar/excelize"
)

const WBSNO_LEN = 17
const WBSNO_POC_LEN = 21

type Col_index_t struct{
	SoNo, Wbs, DbSoNo int
	Product, ProjectName, CustomerNo, CustomerName, ContractNo int
	Classfication int
}

func NewColIndex() Col_index_t{
	return Col_index_t{
		SoNo:-1, Wbs:-1, DbSoNo:-1,
		Product:-1, ProjectName:-1,CustomerNo:-1,CustomerName:-1,ContractNo:-1,
	}
}

func ParseWbsNoValue(val string)string{
	if strings.Index(val, "POC") > 0{
		if len(val) > WBSNO_POC_LEN{
			return val[0:WBSNO_POC_LEN]
		}

		return val
	}
	if len(val) > WBSNO_LEN{
		return val[0:WBSNO_LEN]
	}

	return val
}

func ReadGIS19(xlsx *excelize.File, sheet string)(error, Col_index_t, [][]string){
	colIndex := NewColIndex()
	headerLineNo := 5
	rows := xlsx.GetRows(sheet)
	if len(rows) <= headerLineNo {
		return errors.New("Can not find '" + sheet + "' sheet header"), colIndex, nil
	}

	gisHeader := rows[headerLineNo - 1]


	for index, name := range gisHeader {
		name = strings.TrimSpace(name)
		if name == "SAP Order No.                   Segment  SO Number" {
			colIndex.DbSoNo = index
		}else if name == "CCM--WBS No." {
			colIndex.Wbs = index
		}else if name == "SAP Order No.                 Operation  SO number" {
			colIndex.SoNo = index
		}else if name == "Product" {
			colIndex.Product = index
		}else if name == "Project Name" {
			colIndex.ProjectName = index
		}else if name == "Customer No." {
			colIndex.CustomerNo = index
		}else if name == "Customer Name" {
			colIndex.CustomerName = index
		}else if name == "Contract No." {
			colIndex.ContractNo = index
		}else if strings.Contains(name, "classfication"){
			colIndex.Classfication = index
		}
	}

	if colIndex.SoNo < 0 || colIndex.Wbs < 0 || colIndex.DbSoNo < 0{
		return errors.New("Can not find SO No./WBS/DB SO No. columns"), colIndex, nil
	}

	rows = rows[headerLineNo:]
	var newRows [][]string
	for _, row := range rows {
		if IsEmptyRow(row) {
			continue
		}
		row[colIndex.SoNo] = strings.TrimSpace(row[colIndex.SoNo])
		row[colIndex.Wbs] = ParseWbsNoValue(strings.TrimSpace(row[colIndex.Wbs]))
		row[colIndex.DbSoNo] = strings.TrimSpace(row[colIndex.DbSoNo])

		newRows = append(newRows, row)
	}
	
	return nil, colIndex, newRows
}

func ReadGIS1718(xlsx *excelize.File, sheet, headerLineNoStr string)(error, Col_index_t, [][]string){
		log.Println(sheet + "...")
		headerLineNo, err := strconv.Atoi(headerLineNoStr)
		colIndex := NewColIndex()
		if(err != nil) {
			log.Printf("Header Line No error(%s)\n", headerLineNoStr)
			return err,colIndex,nil
		}

		rows := xlsx.GetRows(sheet)
		if len(rows) <= headerLineNo {
			log.Printf("Can not find '%s' sheet\n", sheet)
			return err,colIndex,nil
		}

		for index, name := range rows[headerLineNo - 1] {
			
			if name == "" {
				continue
			}
			name = strings.ToLower(strings.TrimSpace(name))

			if colIndex.SoNo == -1 && strings.Contains(name, "operation") {
				colIndex.SoNo = index
			} else if colIndex.Wbs == -1 && strings.Contains(name, "wbs") {
				colIndex.Wbs = index
			} else if colIndex.DbSoNo == -1 && strings.Contains(name, "segment") {
				colIndex.DbSoNo = index
			} else if colIndex.Product == -1 && strings.TrimSpace(name) == "product" {
				colIndex.Product = index
			} else if colIndex.ProjectName == -1 && strings.Contains(name, "project name"){
				colIndex.ProjectName = index
			} else if colIndex.CustomerNo == -1 && strings.Contains(name, "customer no") {
				colIndex.CustomerNo = index
			} else if colIndex.CustomerName == -1 && strings.Contains(name, "customer name") {
				colIndex.CustomerName = index
			} else if colIndex.ContractNo == -1 && strings.Contains(name, "contract no") {
				colIndex.ContractNo = index
			}else if colIndex.Classfication == -1 &&  strings.Contains(name, "classfication"){
				colIndex.Classfication = index
			}
		}

		rows = rows[headerLineNo:]
		var newRows [][]string
		for _, row := range rows {
			if IsEmptyRow(row) {
				continue
			}
			if colIndex.SoNo >= 0{
				row[colIndex.SoNo] = strings.TrimSpace(row[colIndex.SoNo])
			}
			if colIndex.DbSoNo >= 0{
				row[colIndex.DbSoNo] = strings.TrimSpace(row[colIndex.DbSoNo])
			}
			if colIndex.Wbs >= 0 {
				row[colIndex.Wbs] = ParseWbsNoValue(strings.TrimSpace(row[colIndex.Wbs]))
			}
			newRows = append(newRows, row)
		}

		return nil,colIndex,newRows
}

func ReadVi19(xlsx *excelize.File, sheet, headerLineNoStr string)(error, Col_index_t, [][]string){
	return ReadGIS1718(xlsx, sheet, headerLineNoStr);
}