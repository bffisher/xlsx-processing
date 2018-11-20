package main

import (
	"log"
	"os"
	"xlsx-processing/copa"
	"xlsx-processing/icbdb"
	"xlsx-processing/unbilledCost"
)

const CONFIG_FILE_PATH = "_conf.xlsx"

func main() {
	task := ""
	if len(os.Args) > 1 {
		task = os.Args[1]
	}

	switch task {
	case "copa":
		execCopa()
	case "icbdb":
		execIcbdb()
	case "UnbilledCost":
		execUnbilledCost()
	default:
		log.Printf("Can not identify the task named '%s'\n", task)
	}
}

func execCopa() {
	err := copa.Exec(CONFIG_FILE_PATH)
	handleResult("copa", err);
}

func execIcbdb() {
	err := icbdb.Exec(CONFIG_FILE_PATH)
	handleResult("icbdb", err);
}

func execUnbilledCost(){
	err := unbilledCost.Exec(CONFIG_FILE_PATH)
	handleResult("unbilled cost", err);
}

func handleResult(name string, err error){
	if err != nil {
		log.Printf("Ocur an error when handling '%s' task.\n", name)
		log.Fatalln(err)
	}

	log.Println("Completed successfully.")
}