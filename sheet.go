package excelparser

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
)

type Worksheet struct {
	XMLName   xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	ColWidths []Col    `xml:"cols>col"`
	Rows      []Row    `xml:"sheetData>row"`
}

func readSheet(sheetName string, xlsxFile *zip.ReadCloser) (*Worksheet, error) {
	// Find the worksheet file in the XLSX archive
	var worksheetFile *zip.File
	for _, f := range xlsxFile.File {
		if f.Name == fmt.Sprintf("xl/worksheets/%s.xml", sheetName) {
			worksheetFile = f
			break
		}
	}
	if worksheetFile == nil {
		return nil, errors.New("worksheet file not found in XLSX archive")
	}

	// Open the worksheet file
	sheet, err := worksheetFile.Open()
	if err != nil {
		return nil, errors.New("failed to open Worksheet file")
	}
	defer sheet.Close()

	// Read the contents of the worksheet
	var buf bytes.Buffer
	if _, err := io.Copy(&buf, sheet); err != nil {
		return nil, errors.New("failed to read Worksheet file")
	}

	var worksheet Worksheet
	err = xml.Unmarshal(buf.Bytes(), &worksheet)
	if err != nil {
		return nil, errors.New("failed to parse from Worksheet")
	}
	return &worksheet, nil
}
