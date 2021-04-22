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

	tgtData := make(map[string]*odata)
	keys := []string{}

	for r := 1; r < srcSheet.MaxRow; r++ {
		key := srcSheet.Cell(r, 0).Value
		theme := srcSheet.Cell(r, 1).Value
		hours, _ := srcSheet.Cell(r, 2).Float()

		var d *odata = nil
		for _, k := range keys {
			if k == key {
				d = tgtData[key]
				break
			}
		}
		if d == nil {
			d = new(odata)
			keys = append(keys, key)
			tgtData[key] = d
		}

		d.key = key
		d.theme = theme
		d.hours += hours
	}

	for _, d := range tgtData {
		log.Printf("\tkey ⇒ %-10v theme ⇒ %-40v hours ⇒ %6.2f", d.key, d.theme, d.hours)
	}

	return xf.Save("./testdata/tgt.xlsx")
}
