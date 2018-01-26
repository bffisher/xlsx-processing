package main

import (
	"log"
	"os"
	"xlsx-processing/copa"
)

const CONFIG_FILE_PATH = "_conf.xlsx"

func main() {
	var what string = ""
	if len(os.Args) > 1 {
		what = os.Args[1]
	}

	switch what {
	case "copa":
		execCopa()
	default:
		log.Fatalln("What do you want to do ?")
	}
}

func execCopa() {
	err := copa.Exec(CONFIG_FILE_PATH)

	if err != nil {
		log.Fatalln("Ocur an error when handling copa.", err)
	}

	log.Println("Completed successfully.")
}
