package copa

import (
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func Exec(confFilePath string) error {
	config, err := handleConfig(confFilePath)
	if err != nil {
		return err
	}

	xlsx, err := excelize.OpenFile(config.filePath)
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
