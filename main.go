package excelparser

import (
	"fmt"
)

func read(fileName string, sheetName string) (map[string][]string, error) {
	xFile, err := readBook(fileName)
	if err != nil {
		return nil, err
	}
	worksheet, err := readSheet(sheetName, xFile)
	if err != nil {
		return nil, err
	}
	sharedString, err := readSharedStrings(xFile)
	if err != nil {
		return nil, err
	}
	data, err := getData(worksheet, sharedString)
	if err != nil {
		return nil, err
	}

	defer xFile.Close()
	return data, nil
}

func main() {
	data, err := read("test1.xlsx", "sheet1")
	fmt.Println(data, err)
}
