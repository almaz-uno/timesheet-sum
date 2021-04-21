package main

import (
	"errors"
	"fmt"
	"log"

	"github.com/tealeg/xlsx"
)

type (
	odata struct {
		key   string
		theme string
		hours float64
	}
)

const tgtSheetName = "summary"

func main() {
	err := processFile("./testdata/Ковров_Максим_Валериевич_2021_01_01_2021_01_31.xlsx")
	if err != nil {
		log.Fatalf("Произошла ошибка разбора файла: %s", err)
	}
	log.Println("Done.")
}

func processFile(xlsxFile string) error {
	xf, err := xlsx.OpenFile(xlsxFile)
	if err != nil {
		return err
	}

	if len(xf.Sheets) == 0 {
		return errors.New("книга пустая")
	}

	srcSheet := xf.Sheets[0]
	tgtSheet, ok := xf.Sheet[tgtSheetName]

	if !ok {
		tgtSheet, err = xf.AddSheet(tgtSheetName)
		if err != nil {
			return fmt.Errorf("не удалось добавить целевой лист %s: %w", tgtSheetName, err)
		}
	}

	_ = srcSheet
	_ = tgtSheet

	return xf.Save("./testdata/tgt.xlsx")
}
