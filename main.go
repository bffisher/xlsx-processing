package main

import (
	"flag"
	"log"
	"xlsx-processing/copa"
)

func main() {
	what := flag.String("what", "", "What do you want to do")
	filePath := flag.String("file", "", "xlsx file full name")
	sheetName := flag.String("sheet", "", "sheet name")
	flag.Parse()

	switch *what {
	case "copa":
		execCopa(filePath, sheetName)
	default:
		log.Fatalln("What do you want to do, -h for help")
	}
}

func execCopa(filePath, sheetName *string) {
	flag.Parse()
	err := copa.Exec(*filePath, *sheetName)

	if err != nil {
		log.Fatalln("Ocur an error when handling copa.", err)
	}
}
