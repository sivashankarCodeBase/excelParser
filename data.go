package excelparser

import "errors"

type Col struct {
	Min   int    `xml:"min,attr"`
	Max   int    `xml:"max,attr"`
	Width int    `xml:"width,attr"`
	Span  string `xml:"span,attr"`
}

type Row struct {
	R     int    `xml:"r,attr"`
	Spans string `xml:"spans,attr"`
	Cells []Cell `xml:"c"`
}

type Cell struct {
	R string `xml:"r,attr"`
	T string `xml:"t,attr,omitempty"`
	V string `xml:"v"`
}

func getData(worksheet *Worksheet, sharedStrings *SharedStrings) (map[string][]string, error) {
	if len(worksheet.Rows) == 0 {
		return nil, errors.New("no records in Worksheet file")
	}

	// Create a map to store cell values for each header
	data := make(map[string][]string)

	// Create a buffered channel to store shared string values
	strch := make(chan string, len(sharedStrings.SharedString))
	for _, stringV := range sharedStrings.SharedString {
		strch <- stringV.T
	}
	close(strch) // Close the channel after sending all values
	headers := []string{}
	// Process worksheet rows
	for n, row := range worksheet.Rows {
		for i, cell := range row.Cells {
			header := ""
			if cell.T == "s" {
				// Retrieve shared string value from channel
				cell.V = <-strch
			}
			// Use the first row to extract headers
			if n == 0 {
				headers = append(headers, cell.V)
			} else {
				// Use headers extracted from the first row to organize cell values
				if i < len(headers) {
					header = headers[i]
					data[header] = append(data[header], cell.V)
				}
			}
		}
	}

	return data, nil
}
