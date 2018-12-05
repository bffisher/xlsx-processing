package main

import (
	"os"
	"log"
	"xlsx-processing/util"
	"xlsx-processing/copa"
	"xlsx-processing/icbdb"
	"xlsx-processing/unbilledCost"
	"xlsx-processing/wip"
)

func main() {
	util.InitEnv()

	task := ""
	if len(os.Args) > 1 {
		task = os.Args[1]
	}

	switch task {
	case "copa":
		err := copa.Exec()
		handleResult("copa", err);
	case "icbdb":
		err := icbdb.Exec()
		handleResult("icbdb", err);
	case "UnbilledCost":
		err := unbilledCost.Exec()
		handleResult("unbilled cost", err);
	case "wip":
		err := wip.Exec()
		handleResult("wip", err);
	default:
		log.Printf("Can not identify the task named '%s'\n", task)
	}
}

func handleResult(name string, err error){
	if err != nil {
		log.Printf("Ocur an error when handling '%s' task.\n", name)
		log.Fatalln(err)
	}

	log.Println("Completed successfully.")
}