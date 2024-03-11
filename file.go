package excelparser

func Extract(xlsxFile *zip.ReadCloser, sheetName string) (map[string][]string, error) {
	worksheet, err := readSheet(sheetName, xlsxFile)
	if err != nil {
		return nil, err
	}
	sharedString, err := readSharedStrings(xlsxFile)
	if err != nil {
		return nil, err
	}
	data, err := getData(worksheet, sharedString)
	if err != nil {
		return nil, err
	}

	defer xlsxFile.Close()
	return data, nil
}
