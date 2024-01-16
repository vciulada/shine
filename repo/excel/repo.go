package excel

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

type SourceType int

const (
	SourceTypeFile SourceType = iota
	SourceTypeFolder
)

type excelData struct {
	sourceType SourceType
	fileName   string
	folderName string
	sheetName  string
	startRow   int
	endRow     int
}

func loadFile(sheetName string, startRow, endRow int) func(string) (chan [][]string, chan error) {
	return func(fileName string) (chan [][]string, chan error) {
		return getFileData(fileName, sheetName, startRow, endRow)
	}
}

func merge(main, secondary [][]string) (result [][]string) {
	return append(main, secondary...)
}

func loadFolderFiles[T any](folder string, loadFile func(string) (chan T, chan error), merge func(T, T) T) (result T) {
	if _, err := os.Stat(folder); os.IsNotExist(err) {
		log.Fatalf("folder %s does not exists: %v\n", folder, err)
	}
	dataChans := []chan T{}
	errorsChan := []chan error{}
	files := []string{}
	if err := filepath.Walk(folder, func(path string, info os.FileInfo, err error) error {
		if !info.IsDir() {
			files = append(files, path)
		}
		return nil
	}); err != nil {
		log.Fatalf("error walking the path %s: %v\n", folder, err)
	}
	for _, path := range files {
		dataChan, errChan := loadFile(path)
		dataChans = append(dataChans, dataChan)
		errorsChan = append(errorsChan, errChan)
	}
	for i := 0; i < len(files); i++ {
		select {
		case data := <-dataChans[i]:
			result = merge(result, data)
			continue
		case err := <-errorsChan[i]:
			log.Fatal(err)
		}
	}
	return
}

func getFileData(fileName, sheetName string, startRow, endRow int) (chan [][]string, chan error) {
	result := make(chan [][]string)
	errChan := make(chan error)
	go func() {
		f, err := excelize.OpenFile(fileName)
		if err != nil {
			errChan <- err
		}
		defer func() {
			// Close the spreadsheet.
			if err := f.Close(); err != nil {
				fmt.Println(err)
			}
		}()
		// Get all the rows in the Sheet1.
		rows, err := f.GetRows(sheetName)
		if err != nil {
			errChan <- err
		}
		if endRow > 0 {
			result <- rows[startRow:endRow]
		} else {
			result <- rows[startRow:]
		}
	}()
	return result, errChan
}

func (e excelData) GetData() [][]string {
	switch e.sourceType {
	case SourceTypeFile:
		result, errChan := getFileData(e.fileName, e.sheetName, e.startRow, e.endRow)
		select {
		case data := <-result:
			return data
		case err := <-errChan:
			log.Fatal(err)
		}
	case SourceTypeFolder:
		return loadFolderFiles[[][]string](e.folderName, loadFile(e.sheetName, e.startRow, e.endRow), merge)
	}
	return [][]string{}
}

func New(sourceType SourceType, name, sheetName string, startRow int, endRow int) excelData {
	switch sourceType {
	case SourceTypeFile:
		return excelData{sourceType: sourceType, fileName: name, sheetName: sheetName, startRow: startRow, endRow: endRow}
	case SourceTypeFolder:
		return excelData{sourceType: sourceType, folderName: name, sheetName: sheetName, startRow: startRow, endRow: endRow}
	default:
		log.Fatalf("unknown source type:%d", sourceType)
	}
	return excelData{}
}
