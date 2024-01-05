package excel

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

type excelData struct {
	fileName  string
	sheetName string
	startRow  int
}

func (e excelData) GetData() [][]string {
	f, err := excelize.OpenFile(e.fileName)
	if err != nil {
		log.Fatal(err)
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows(e.sheetName)
	if err != nil {
		log.Fatal(err)
	}
	return rows[e.startRow:]
}

func New(fileName, sheetName string, startRow int) excelData {
	return excelData{fileName: fileName, sheetName: sheetName, startRow: startRow}
}
