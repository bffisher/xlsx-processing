package util

import (
	"errors"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

//row: row index, col: column index, them start from 0
func Axis(row, col int) string {
	row++
	if col < 26 {
		return string('A'+col) + strconv.Itoa(row)
	}

	return string('A'+(col)/26-1) + string('A'+col%26) + strconv.Itoa(row)
}

//from string like "{code}:{name}" to extract code, and trim spaces
func ExtractCode(value string) string {
	return strings.TrimSpace(strings.Split(value, ":")[0])
}

//Split string like "{code}:{name}", return code and name, and trim spaces
func SplitCodeName(value string) (code, name string) {
	array := strings.Split(value, ":")
	code = strings.TrimSpace(array[0])
	if len(array) > 1 {
		name = strings.TrimSpace(array[1])
	}
	return
}

func ReadConfig(filePath string) ([][]string, error) {
	xlsx, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}

	rows := xlsx.GetRows(xlsx.GetSheetName(1))
	if len(rows) < 2 {
		return nil, errors.New("Read config info failed. No data.")
	}

	return rows[1:], nil
}

func CopyRowToMap(dic map[string]string, row []string, keyIdx, valIdx int) {
	if row[keyIdx] != "" {
		dic[row[keyIdx]] = row[valIdx]
	}
}

func CopyRowToMapWithFunc(dic map[string]string, row []string, keyIdx, valIdx int, handle func(string, string) (string, string)) {
	if row[keyIdx] != "" {
		key, value := handle(row[keyIdx], row[valIdx])
		dic[key] = value
	}
}

func ConvColNameToIdx(row []string, colNames map[string]string, colIdxs map[string]int, keyFilter func(key string) bool) {
	for index, col := range row {
		col := strings.TrimSpace(col)
		for key, name := range colNames {
			if name == col && (keyFilter == nil || keyFilter(key)) {
				colIdxs[key] = index
				break
			}
		}
	}
}

func CheckColNameIdx(colNames map[string]string, colIdxs map[string]int) error {
	for key, name := range colNames {
		if _, ok := colIdxs[key]; !ok {
			return errors.New("Can not find '" + name + "' column")
		}
	}
	return nil
}
