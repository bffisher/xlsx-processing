package unbilledCost
import (
	//"errors"
	"log"
	//"strconv"
	//"strings"
	//"xlsx-processing/util"

	"github.com/360EntSecGroup-Skylar/excelize"
)
type data_t struct {
	conf *config_t
	ucXlsx, ioXlsx, fgXlsx, sgXlsx *excelize.File
}
func Exec(confFile string) error {
	var err error

	data:=&data_t{}
	log.Print("Reading... ")
	
	data.conf, err = readConfig(confFile)
	if err != nil {
		return err
	}
	data.ucXlsx, err = excelize.OpenFile(data.conf.files["UC_FILE"])
	if err != nil {
		return err
	}
	data.ioXlsx, err = excelize.OpenFile(data.conf.files["IO_FILE"])
	if err != nil {
		return err
	}
	data.fgXlsx, err = excelize.OpenFile(data.conf.files["FG_FILE"])
	if err != nil {
		return err
	}
	data.sgXlsx, err = excelize.OpenFile(data.conf.files["SG_FILE"])
	if err != nil {
		return err
	}
	
	log.Println("OK!")
	return nil
}