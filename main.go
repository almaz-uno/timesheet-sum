package main

import (
	"bytes"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"strings"

	tgbotapi "github.com/go-telegram-bot-api/telegram-bot-api/v5"
	"github.com/kr/pretty"
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
	apiToken := os.Getenv("TELEGRAM_APITOKEN")
	if len(apiToken) == 0 {
		log.Fatalf("Задайте переменную окружения TELEGRAM_APITOKEN")
	}
	bot, err := tgbotapi.NewBotAPI(apiToken)
	if err != nil {
		log.Fatalf("Ошибка открытия Telegram API: %s", err.Error())
	}

	bot.Debug = true
	ucfg := tgbotapi.UpdateConfig{
		Offset:  0,
		Timeout: 30,
	}
	updates := bot.GetUpdatesChan(ucfg)

	for u := range updates {
		if u.Message == nil {
			continue
		}
		pretty.Println("update:", u)

		replyText := func(text string) {
			msg := tgbotapi.NewMessage(u.Message.Chat.ID, text)
			msg.ReplyToMessageID = u.Message.MessageID
			if _, e2 := bot.Send(msg); e2 != nil {
				log.Printf("Ошибка отправки ответа: %s", e2)
			}
		}

		if u.Message.Document == nil {
			replyText("Пришлите xlsx-документ выгрузки из JIRA")
			continue
		}

		fi, err := bot.GetFile(tgbotapi.FileConfig{FileID: u.Message.Document.FileID})
		if err != nil {
			replyText("Произошла ошибка получения файла")
			log.Println("Ошибка получения информации о файле:", err)
			continue
		}
		fileURL := fmt.Sprintf("https://api.telegram.org/file/bot%s/%s", apiToken, fi.FilePath)

		res, err := http.Get(fileURL)
		if err != nil {
			replyText("Произошла ошибка получения файла" + fi.FilePath)
			log.Printf("Ошибка получения файла %s: %s", fi.FilePath, err)
			continue
		}
		defer res.Body.Close()

		bb, err := processXlsx(res.Body)
		if err != nil {
			replyText("Произошла ошибка обработки файла: " + err.Error())
			log.Printf("Произошла ошибка обработки файла %s: %s", fi.FilePath, err)
			continue
		}

		msgDoc := tgbotapi.NewDocument(u.Message.Chat.ID, tgbotapi.FileBytes{
			Name:  u.Message.Document.FileName,
			Bytes: bb,
		})

		msgDoc.ReplyToMessageID = u.Message.MessageID

		if _, e2 := bot.Send(msgDoc); e2 != nil {
			log.Printf("Ошибка отправки ответа: %s", e2)
		}

	}

	log.Println("Done.")
}

func processXlsx(srcXlsx io.Reader) ([]byte, error) {
	bb, err := ioutil.ReadAll(srcXlsx)
	if err != nil {
		return nil, err
	}
	xf, err := xlsx.OpenBinary(bb)
	if err != nil {
		return nil, err
	}

	if len(xf.Sheets) == 0 {
		return nil, errors.New("книга пустая")
	}

	srcSheet := xf.Sheets[0]
	tgtSheet, ok := xf.Sheet[tgtSheetName]

	if !ok {
		tgtSheet, err = xf.AddSheet(tgtSheetName)
		if err != nil {
			return nil, fmt.Errorf("не удалось добавить целевой лист %s: %w", tgtSheetName, err)
		}
	}

	tgtData, keys := extract(srcSheet)

	fmt.Println(toString(tgtData, keys))

	addData(tgtSheet, tgtData, keys)

	buff := &bytes.Buffer{}

	err = xf.Write(buff)
	if err != nil {
		return nil, err
	}

	return buff.Bytes(), nil
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
