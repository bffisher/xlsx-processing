package util

import (
	"strconv"
	"strings"
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
