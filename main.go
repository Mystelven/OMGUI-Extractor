package main

import (
	"log"

	"bytes"
	"encoding/csv"
	"fmt"
	"io"
	"io/ioutil"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/harry1453/go-common-file-dialog/cfd"
	"github.com/harry1453/go-common-file-dialog/cfdutil"
)

const (
	LOGO string = ""
)

const (
	NB_LINES_PER_DAY int    = 1441
	FIRST_SHEET      string = "Semaine"
)

var purple int = 0
var blue int = 0

func LinesBytesCount(s []byte) int {
	nl := []byte{'\n'}
	n := bytes.Count(s, nl)
	if len(s) > 0 && !bytes.HasSuffix(s, nl) {
		n++
	}
	return n
}

func SetLegend(f *excelize.File, sheets []string) *excelize.File {

	for _, sheet := range sheets {

		f.SetCellValue(sheet, "A1", "Date et heure")
		f.SetCellValue(sheet, "B1", "Minutes sédentaires")
		f.SetCellValue(sheet, "C1", "Minutes AP faible intensitée")
		f.SetCellValue(sheet, "D1", "Minutes AP intensité modérée")
		f.SetCellValue(sheet, "E1", "Minutes AP forte intensitée")

		f.SetCellValue(sheet, "H1", "Sédentaires (total)")
		f.SetCellValue(sheet, "I1", "Faible intensitée (total)")
		f.SetCellValue(sheet, "J1", "Intensité modérée (total)")
		f.SetCellValue(sheet, "K1", "Forte intensitée (total)")
	}

	return f
}

func InitializeSheets(f *excelize.File, nbPages int) (*excelize.File, []string) {

	sheets := []string{}

	f.NewSheet(FIRST_SHEET)
	f.DeleteSheet("Sheet1")

	for i := 1; i <= ((nbPages / NB_LINES_PER_DAY) + 1); i++ {
		f.NewSheet("J" + strconv.Itoa(i))
		sheets = append(sheets, "J"+strconv.Itoa(i))
	}

	sheets = append(sheets, FIRST_SHEET)

	return f, sheets
}

func SetAllValuesInSheet(f *excelize.File, sheet string, record []string, i int) (int, int, int, int) {

	r1, _ := strconv.Atoi(record[1])
	r2, _ := strconv.Atoi(record[2])
	r3, _ := strconv.Atoi(record[3])
	r4, _ := strconv.Atoi(record[4])

	if r1 == 1 {
		f.SetCellStyle(sheet, "B"+strconv.Itoa(i+1), "B"+strconv.Itoa(i+1), blue)
	}

	if r2 == 1 {
		f.SetCellStyle(sheet, "C"+strconv.Itoa(i+1), "C"+strconv.Itoa(i+1), blue)
	}

	if r3 == 1 {
		f.SetCellStyle(sheet, "D"+strconv.Itoa(i+1), "D"+strconv.Itoa(i+1), blue)
	}

	if r4 == 1 {
		f.SetCellStyle(sheet, "E"+strconv.Itoa(i+1), "E"+strconv.Itoa(i+1), blue)
	}

	if r1+r2+r3+r4 != 1 {
		fmt.Println("We have a problem at line: " + strconv.Itoa(i) + " in sheet: " + sheet)
	} else {
		fmt.Println("Sheet [" + sheet + "] -- line: " + strconv.Itoa(i) + ": OK")
	}

	f.SetCellValue(sheet, "A"+strconv.Itoa(i+1), record[0])
	f.SetCellValue(sheet, "B"+strconv.Itoa(i+1), r1)
	f.SetCellValue(sheet, "C"+strconv.Itoa(i+1), r2)
	f.SetCellValue(sheet, "D"+strconv.Itoa(i+1), r3)
	f.SetCellValue(sheet, "E"+strconv.Itoa(i+1), r4)

	return r1, r2, r3, r4

}

func SetFinalCount(f *excelize.File, sheet string, nbSedentary int, nbLights int, nbModerate int, nbVigorous int) {

	f.SetCellValue(sheet, "H2", nbSedentary)
	f.SetCellValue(sheet, "I2", nbLights)
	f.SetCellValue(sheet, "J2", nbModerate)
	f.SetCellValue(sheet, "K2", nbVigorous)

	f.SetCellValue(sheet, "H3", nbSedentary+nbLights+nbModerate+nbVigorous)
}

func ColorOnes(f *excelize.File, sheet string, i int, j int) {

	if j == 3 {

		f.SetCellStyle(sheet, "D"+strconv.Itoa(i+2), "D"+strconv.Itoa(i+2), purple)
		f.SetCellStyle(sheet, "D"+strconv.Itoa(i+1), "D"+strconv.Itoa(i+1), purple)
		f.SetCellStyle(sheet, "D"+strconv.Itoa(i), "D"+strconv.Itoa(i), purple)
		f.SetCellStyle(sheet, "D"+strconv.Itoa(i-1), "D"+strconv.Itoa(i-1), purple)
		f.SetCellStyle(sheet, "D"+strconv.Itoa(i-2), "D"+strconv.Itoa(i-2), purple)
		f.SetCellStyle(sheet, "D"+strconv.Itoa(i-3), "D"+strconv.Itoa(i-3), purple)
		f.SetCellStyle(sheet, "D"+strconv.Itoa(i-4), "D"+strconv.Itoa(i-4), purple)
		f.SetCellStyle(sheet, "D"+strconv.Itoa(i-5), "D"+strconv.Itoa(i-5), purple)
		f.SetCellStyle(sheet, "D"+strconv.Itoa(i-6), "D"+strconv.Itoa(i-6), purple)
		f.SetCellStyle(sheet, "D"+strconv.Itoa(i-7), "D"+strconv.Itoa(i-7), purple)
		// f.SetCellStyle(sheet, "D"+strconv.Itoa(i-8), "D"+strconv.Itoa(i-8), purple)
	}

	if j == 4 {

		f.SetCellStyle(sheet, "E"+strconv.Itoa(i+2), "E"+strconv.Itoa(i+2), purple)
		f.SetCellStyle(sheet, "E"+strconv.Itoa(i+1), "E"+strconv.Itoa(i+1), purple)
		f.SetCellStyle(sheet, "E"+strconv.Itoa(i), "E"+strconv.Itoa(i), purple)
		f.SetCellStyle(sheet, "E"+strconv.Itoa(i-1), "E"+strconv.Itoa(i-1), purple)
		f.SetCellStyle(sheet, "E"+strconv.Itoa(i-2), "E"+strconv.Itoa(i-2), purple)
		f.SetCellStyle(sheet, "E"+strconv.Itoa(i-3), "E"+strconv.Itoa(i-3), purple)
		f.SetCellStyle(sheet, "E"+strconv.Itoa(i-4), "E"+strconv.Itoa(i-4), purple)
		f.SetCellStyle(sheet, "E"+strconv.Itoa(i-5), "E"+strconv.Itoa(i-5), purple)
		f.SetCellStyle(sheet, "E"+strconv.Itoa(i-6), "E"+strconv.Itoa(i-6), purple)
		f.SetCellStyle(sheet, "E"+strconv.Itoa(i-7), "E"+strconv.Itoa(i-7), purple)
		// f.SetCellStyle(sheet, "E"+strconv.Itoa(i-8), "E"+strconv.Itoa(i-8), purple)
	}
}

func CheckIfColorSuccessiveOnes(f *excelize.File, sheet string, records [][]string) {

	nbSuccessiveModerate := 0
	nbSuccessiveVigourus := 0

	nbBlockModerate := 0
	nbBlockVigourus := 0

	lastCounterModerate := false
	lastCounterVigourous := false

	for i := 0; i < len(records); i++ {

		r1, _ := strconv.Atoi(records[i][3])
		r2, _ := strconv.Atoi(records[i][4])

		if r1 == 0 {

			lastCounterModerate = false
			nbSuccessiveModerate = 0

		} else {
			nbSuccessiveModerate++
		}

		if nbSuccessiveModerate >= 10 {

			if !lastCounterModerate {
				nbBlockModerate++
				lastCounterModerate = true
			}
			ColorOnes(f, sheet, i, 3)
		}

		if r2 == 0 {
			lastCounterVigourous = false
			nbSuccessiveVigourus = 0

		} else {
			nbSuccessiveVigourus++
		}

		if nbSuccessiveVigourus >= 10 {
			if !lastCounterVigourous {
				nbBlockVigourus++
				lastCounterVigourous = true
			}
			ColorOnes(f, sheet, i, 4)
		}
	}

	f.SetCellValue(sheet, "H6", "#Blocs (Intensitée Modérée)")
	f.SetCellValue(sheet, "H7", "#Blocs (Forte Intensitée)")

	f.SetCellValue(sheet, "I6", nbBlockModerate)
	f.SetCellValue(sheet, "I7", nbBlockVigourus)
}

func CountNbCellPurples(f *excelize.File, sheet string, nbLines int) {

	nbModeratePurple := 0
	nbVigourousPurple := 0
	var err error
	defer func() {
		if r := recover(); r != nil {
			err = r.(error)
		}
	}()
	for i := 0; i < nbLines; i++ {

		styleId := f.GetCellStyle(sheet, "D"+strconv.Itoa(i))
		if styleId == purple {
			nbModeratePurple++
		}

		styleId = f.GetCellStyle(sheet, "E"+strconv.Itoa(i))
		if styleId == purple {
			nbVigourousPurple++
		}
	}

	f.SetCellValue(sheet, "K6", "Minutes AP intensité modérée par bloc min 10 mins")
	f.SetCellValue(sheet, "K7", "Minutes AP forte intensité par bloc min 10 mins")

	f.SetCellValue(sheet, "L6", nbModeratePurple)
	f.SetCellValue(sheet, "L7", nbVigourousPurple)
}

func SplitCSVIntoMultiplesSheetsExcel(filename string) error {

	f := excelize.NewFile()

	p, _ := f.NewStyle(`{"fill":{"type":"pattern","color":["#884EA0"],"pattern":1}}`)
	b, _ := f.NewStyle(`{"fill":{"type":"pattern","color":["#2471A3"],"pattern":1}}`)

	purple = p
	blue = b

	dat, err := ioutil.ReadFile(filename)
	if err != nil {
		return err
	}

	rAll := csv.NewReader(bytes.NewReader(dat))
	records, _ := rAll.ReadAll()

	f, sheets := InitializeSheets(f, LinesBytesCount(dat))
	f = SetLegend(f, sheets)

	r := csv.NewReader(bytes.NewReader(dat))
	r.Read() // We do not need the header, we know it.

	i := 1
	j := 0

	nbSedentary := 0
	nbLights := 0
	nbModerate := 0
	nbVigourus := 0

	dataRecord := make([][]string, 0)

	for {

		record, err := r.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		dataRecord = append(dataRecord, record)

		r1, r2, r3, r4 := SetAllValuesInSheet(f, sheets[j], record, i)
		nbSedentary += r1
		nbLights += r2
		nbModerate += r3
		nbVigourus += r4

		i++
		if i == NB_LINES_PER_DAY {

			SetFinalCount(f, sheets[j], nbSedentary, nbLights, nbModerate, nbVigourus)

			CheckIfColorSuccessiveOnes(f, sheets[j], dataRecord)

			dataRecord = make([][]string, 0)

			CountNbCellPurples(f, sheets[j], NB_LINES_PER_DAY)

			nbSedentary = 0
			nbLights = 0
			nbModerate = 0
			nbVigourus = 0

			i = 1
			j++
		}
	}

	r = csv.NewReader(bytes.NewReader(dat))
	r.Read() // We do not need the header, we know it.

	nbSedentary = 0
	nbLights = 0
	nbModerate = 0
	nbVigourus = 0

	for {
		record, err := r.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		r1, r2, r3, r4 := SetAllValuesInSheet(f, FIRST_SHEET, record, i)
		nbSedentary += r1
		nbLights += r2
		nbModerate += r3
		nbVigourus += r4
		i++
	}

	SetFinalCount(f, FIRST_SHEET, nbSedentary, nbLights, nbModerate, nbVigourus)

	CheckIfColorSuccessiveOnes(f, FIRST_SHEET, records)

	CountNbCellPurples(f, sheets[j], len(records))

	return f.SaveAs(filename + ".xlsx")
}

func main() {
	result, err := cfdutil.ShowOpenFileDialog(cfd.DialogConfig{
		Title: "OMGUI-Extractor",
		Role:  "OpenFileExample",
		FileFilters: []cfd.FileFilter{
			{
				DisplayName: "CSV Files (*.csv)",
				Pattern:     "*.csv",
			},
		},
		SelectedFileFilterIndex: 2,
		FileName:                "",
		DefaultExtension:        "csv",
	})
	if err != nil {
		log.Fatal(err)
	}

	log.Printf("Chosen file: %s\n", result)

	SplitCSVIntoMultiplesSheetsExcel(result)
}
