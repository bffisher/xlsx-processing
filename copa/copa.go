package copa

import (
	"strings"
	"xlsx-processing/util"
	"github.com/360EntSecGroup-Skylar/excelize"
)

func Exec() error {
	config, err := handleConfig(util.Env().ConfigFullName)
	if err != nil {
		return err
	}

	xlsx, err := excelize.OpenFile(util.Env().FilePath)
	if err != nil {
		return err
	}

	header, err := handleHeader(xlsx, &config)
	if err != nil {
		return err
	}

	handleBody(xlsx, &config, &header)

	return xlsx.SaveAs(strings.Replace(config.filePath, ".xlsx", "_res.xlsx", -1))
}
