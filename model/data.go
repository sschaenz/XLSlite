package model

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// DataEntry is a struct that stores the column names of the Excel file.
type DataEntry struct {
	Topic          string
	ControlID      string
	Description    string
	Condition      string
	Domain         string
	Scope          string
	Implementation string
	Finding        string
	Proof          string
	MaturityLevel  string
	Recommendation string
	Note           string
}

// ExcelHeader is a struct that stores the column in A1-notation of the Excel file.
type ExcelHeader struct {
	Topic          string
	ControlID      string
	Description    string
	Condition      string
	Domain         string
	Scope          string
	Implementation string
	Finding        string
	Proof          string
	MaturityLevel  string
	Recommendation string
	Note           string
}

// ExcelReader is a struct that reads data from an Excel file.
type ExcelReader struct {
	f         *excelize.File
	sheet     string
	headerRow int
	header    ExcelHeader
}

// NewExcelReader creates a new ExcelReader struct.
func NewExcelReader(source, sheet string, headerRow int) (*ExcelReader, error) {
	f, err := excelize.OpenFile(source)
	if err != nil {
		return nil, fmt.Errorf("fehler beim Öffnen der Datei: %v", err)
	}

	reader := &ExcelReader{
		f:         f,
		sheet:     sheet,
		headerRow: headerRow,
	}

	err = reader.readHeader()
	if err != nil {
		return nil, err
	}

	return reader, nil
}

// readHeader reads the header row and stores the column names in the ExcelReader struct.
func (r *ExcelReader) readHeader() error {
	headerNames := []string{
		"topic", "control id", "beschreibung", "anforderung", "domain", "audit scope", "beschreibung der aktuellen umsetzung", "feststellung", "nachweis", "erfüllungsgrad", "empfehlung", "bemerkung",
	}

	row, err := r.f.GetCols(r.sheet)
	if err != nil {
		return fmt.Errorf("fehler beim Abrufen der Header-Zeile: %v", err)
	}

	for i, cell := range row[r.headerRow-1] {
		columnName := strings.ToLower(cell)
		column, _ := excelize.ColumnNumberToName(i + 1)
		for _, headerName := range headerNames {
			if columnName == headerName {
				switch headerName {
				case "topic":
					r.header.Topic = column
				case "control id":
					r.header.ControlID = column
				case "beschreibung":
					r.header.Description = column
				case "anforderung":
					r.header.Condition = column
				case "domain":
					r.header.Domain = column
				case "audit scope":
					r.header.Scope = column
				case "beschreibung der aktuellen umsetzung":
					r.header.Implementation = column
				case "feststellung":
					r.header.Finding = column
				case "nachweis":
					r.header.Proof = column
				case "erfüllungsgrad":
					r.header.MaturityLevel = column
				case "empfehlung":
					r.header.Recommendation = column
				case "bemerkung":
					r.header.Note = column
				}
			}
		}
	}

	return nil
}

// ReadData liest die Daten aus der Excel-Datei aus.
func (r *ExcelReader) ReadData(contentRow, numOfRows int) ([]DataEntry, error) {
	var data []DataEntry

	rows, err := r.f.Rows(r.sheet)
	if err != nil {
		return nil, fmt.Errorf("fehler beim Abrufen der Zeilen: %v", err)
	}

	rowIndex := 1
	for rows.Next() {
		if rowIndex < contentRow {
			rowIndex++
			continue
		}

		if numOfRows != 0 && rowIndex >= contentRow+numOfRows {
			break
		}

		entry := &DataEntry{}

		entry.Topic, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.ControlID, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.Description, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.Condition, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.Domain, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.Scope, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.Implementation, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.Finding, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.Proof, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.MaturityLevel, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.Recommendation, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))
		entry.Note, _ = r.f.GetCellValue(r.sheet, fmt.Sprintf("%s%d", r.header.Topic, rowIndex))

		data = append(data, *entry)

		rowIndex++
	}

	return data, nil
}

// ExcelToDB reads the data from the excel file and writes it to the database
func ExcelToDB(source string, sheet string, headerRow int, contentRow int, numRows int, output string) (*[]DataEntry, error) {
	// read Excel file
	reader, err := NewExcelReader(source, sheet, headerRow)
	if err != nil {
		return nil, err
	}
	// read header
	err = reader.readHeader()
	// read data from Excel file
	data, err := reader.ReadData(contentRow, numRows)
	if err != nil {
		return nil, err
	}
	// write data to database
	// TODO: implement
	return &data, nil
}
