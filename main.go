package main

import (
	"log"
	"os"
	"xlsx-processing/copa"
)

const CONFIG_FILE_PATH = "_conf.xlsx"

func main() {
	var task string = ""
	if len(os.Args) > 1 {
		task = os.Args[1]
	}

	switch task {
	case "copa":
		execCopa()
	default:
		log.Printf("Can not identify the task named '%s'\n", task)
	}
}

func execCopa() {
	err := copa.Exec(CONFIG_FILE_PATH)

	if err != nil {
		log.Println("Ocur an error when handling 'copa' task.")
		log.Fatalln(err)
	}

	log.Println("Completed successfully.")
}
