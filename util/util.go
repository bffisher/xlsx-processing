package util

import (
	"strconv"
)

func Axis(row, col int) string {
	row++
	if col < 26 {
		return string('A'+col) + strconv.Itoa(row)
	}

	return string('A'+(col)/26-1) + string('A'+col%26) + strconv.Itoa(row)
}
