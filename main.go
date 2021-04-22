package main

import (
	"errors"
	"fmt"
	"log"
	"strings"

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

var fields = []string{"key", "theme", "hours", "days"}

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

	tgtData, keys := extract(srcSheet)

	fmt.Println(toString(tgtData, keys))

	addData(tgtSheet, tgtData, keys)

	return xf.Save("./testdata/tgt.xlsx")
}

func extract(srcSheet *xlsx.Sheet) (map[string]*odata, []string) {
	tgtData := map[string]*odata{}
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

	return tgtData, keys
}

func toString(tgtData map[string]*odata, keys []string) string {
	b := &strings.Builder{}
	for _, k := range keys {
		d := tgtData[k]
		fmt.Fprintf(b, "\tkey ⇒ %-10v theme ⇒ %-40v hours ⇒ %6.2f\n", d.key, d.theme, d.hours)
	}
	return b.String()
}

func addData(tgtSheet *xlsx.Sheet, tgtData map[string]*odata, keys []string) {
	row := tgtSheet.AddRow()
	for _, f := range fields {
		cell := row.AddCell()
		st := cell.GetStyle()
		st.Font.Bold = true
		st.Alignment.Horizontal = "center"
		st.ApplyAlignment = true
		st.ApplyFont = true
		cell.SetString(f)
	}

	for _, k := range keys {
		d := tgtData[k]
		row = tgtSheet.AddRow()

		row.AddCell().SetString(d.key)
		row.AddCell().SetString(d.theme)

		row.AddCell().SetFloat(d.hours)
		row.AddCell().SetFloat(d.hours / 8.0)

	}
}
