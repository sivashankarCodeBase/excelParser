package excelparser

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
)

// SharedStrings represents the structure of the shared strings XML.
type SharedStrings struct {
	XMLName      xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Count        int      `xml:"count,attr"`
	UniqueCount  int      `xml:"uniqueCount,attr"`
	SharedString []SI     `xml:"si"`
}

// SI represents the shared string item.
type SI struct {
	T string `xml:"t"`
}

func readSharedStrings(xlsxFile *zip.ReadCloser) (*SharedStrings, error) {
	// Find the worksheet file in the XLSX archive
	var worksheetFile *zip.File
	for _, f := range xlsxFile.File {
		if f.Name == "xl/sharedStrings.xml" {
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

	var sharedStrings SharedStrings
	err = xml.Unmarshal(buf.Bytes(), &sharedStrings)
	if err != nil {
		return nil, errors.New("failed to parse from Worksheet")
	}

	return &sharedStrings, nil
}
