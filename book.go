package excelparser

import (
	"archive/zip"
	"errors"
)

func Read(fileName string) (zip.ReadCloser, error) {
	// Open the XLSX file
	xFile, err := zip.OpenReader(fileName)
	if err != nil {
		return nil, errors.New("failed to open the file")
	}
	return xFile, nil
}
