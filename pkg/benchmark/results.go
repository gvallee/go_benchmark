//
// Copyright (c) 2025, NVIDIA CORPORATION. All rights reserved.
//
// See LICENSE.txt for license information
//

package benchmark

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/gvallee/go_util/pkg/notation"
)

type DataPoint struct {
	Size  float64
	Value float64
}

type Result struct {
	DataPoints []*DataPoint
}

type Results struct {
	Result []*Result
}

// SpreadsheetData represents the data to be saved in a spreadsheet as well as
// all the information about how to save it (e.g., which sheet)
type SpreadsheetData struct {
	// SheetStart is the unique sheet ID where we will start saving the data (most of the time a single sheet)
	SheetStart int

	// Data is the OSU data to save in the spreadsheet
	Data *Results

	// Labels is the ordered list of labels associated to the OSU data
	Labels []string
}

// SpreadsheetMetadata is the metadata associated to the data
type SpreadsheetMetadata struct {
	// SheetID is the unique sheet ID where all the metadata will be saved (single sheet)
	SheetID int

	// Timestamp is the time stamp in string version associated with the entire experiment
	Timestamp string

	// Metadata content, one element of the slice per line in the spreadsheet (only one column for now)
	Content []string
}

func addValuesToExcel(excelFile *excelize.File, sheetID string, lineStart int, col int, datapoints []*DataPoint) error {
	colID := notation.IntToAA(col)
	lineID := lineStart
	for _, d := range datapoints {
		// Find the correct line where to put the data
		for {
			dataSizeStr := excelFile.GetCellValue(sheetID, fmt.Sprintf("A%d", lineID))
			dataSize, err := strconv.ParseFloat(dataSizeStr, 64)
			if err != nil {
				return fmt.Errorf("unable to parse %s: %w", dataSizeStr, err)
			}
			if dataSize == d.Size {
				break
			}
			lineID++
		}
		excelFile.SetCellValue(sheetID, fmt.Sprintf("%s%d", colID, lineID), d.Value)
		lineID++
	}
	return nil
}

func prepSheet(excelFile *excelize.File, sheetNum int) (string, error) {
	if excelFile == nil {
		return "", fmt.Errorf("undefined excelFile object")
	}

	sheetID := fmt.Sprintf("Sheet%d", sheetNum)

	if sheetID != "Sheet1" {
		excelFile.NewSheet(sheetID)
	}
	return sheetID, nil
}

func addMetadataToSpreadsheet(excelFile *excelize.File, spreadsheetMetadata *SpreadsheetMetadata) error {
	if excelFile == nil {
		return fmt.Errorf("undefined excelFile object")
	}

	if spreadsheetMetadata == nil {
		return fmt.Errorf("undefined metadata")
	}

	sheetID, err := prepSheet(excelFile, spreadsheetMetadata.SheetID)
	if err != nil {
		return fmt.Errorf("prepSheet() failed: %w", err)
	}

	lineID := 1 // 1-indexed to match Excel semantics
	col := 0    // 0-indexed so it can be used with IntToAA

	// Timestamp
	excelFile.SetCellValue(sheetID, fmt.Sprintf("%s%d", notation.IntToAA(col), lineID), spreadsheetMetadata.Timestamp)
	lineID++

	// Metadata content passed in by the user
	for _, line := range spreadsheetMetadata.Content {
		excelFile.SetCellValue(sheetID, fmt.Sprintf("%s%d", notation.IntToAA(col), lineID), line)
		lineID++
	}

	return nil
}

func addDataToSpreadsheet(excelFile *excelize.File, spreadsheetData *SpreadsheetData) error {
	if excelFile == nil {
		return fmt.Errorf("undefined excelFile object")
	}

	if spreadsheetData == nil {
		return fmt.Errorf("undefined data")
	}

	sheetID, err := prepSheet(excelFile, spreadsheetData.SheetStart)
	if err != nil {
		return fmt.Errorf("prepSheet() failed: %w", err)
	}

	// Add the labels
	lineID := 1 // 1-indexed to match Excel semantics
	col := 1    // 0-indexed so it can be used with IntToAA
	for _, label := range spreadsheetData.Labels {
		excelFile.SetCellValue(sheetID, fmt.Sprintf("%s%d", notation.IntToAA(col), lineID), label)
		col++
	}

	// Add the message sizes into the first column
	lineID = 2 // 1-indexed
	for _, dp := range spreadsheetData.Data.Result[0].DataPoints {
		excelFile.SetCellValue(sheetID, fmt.Sprintf("A%d", lineID), dp.Size)
		lineID++
	}

	// Add the values
	col = 1    // 0-indexed so it can be used with IntToAA
	lineID = 2 // 1-indexed
	for _, d := range spreadsheetData.Data.Result {
		err := addValuesToExcel(excelFile, sheetID, lineID, col, d.DataPoints)
		if err != nil {
			return fmt.Errorf("addValuesToExcel() failed: %w", err)
		}
		col++
	}
	return nil
}

// Excelize creates a very simple spreadsheet with only the raw OSU data
func Excelize(excelFilePath string, results *Results) error {
	if results == nil {
		return fmt.Errorf("undefined data")
	}

	excelFile := excelize.NewFile()

	// Add the message sizes into the first column
	lineID := 1 // 1-indexed to match Excel semantics
	for _, dp := range results.Result[0].DataPoints {
		excelFile.SetCellValue("Sheet1", fmt.Sprintf("A%d", lineID), dp.Size)
		lineID++
	}

	// Add the values
	col := 1   // 0-indexed so it can be used with IntToAA
	lineID = 1 // 1-indexed to match Excel semantics
	for _, d := range results.Result {
		err := addValuesToExcel(excelFile, "Sheet1", lineID, col, d.DataPoints)
		if err != nil {
			return fmt.Errorf("addValuesToExcel() failed: %w", err)
		}
		col++
	}

	err := excelFile.SaveAs(excelFilePath)
	if err != nil {
		return err
	}

	return nil
}

// ExcelizeWithLabels create a MSExcel spreadsheet with all the data and metadata passed in.
// The metadata is saved on a separate sheet and meant to capture all the necessary
// details to understand the data and how it was gathered.
// The data includes the OSU data and the corresponding labels associated to the data.
// The order of the labels is assumed to be the same than the order of the data.
// All references to sheets is 1-based indexed.
func ExcelizeWithLabels(spreadsheetMetadata *SpreadsheetMetadata, spreadsheetData *SpreadsheetData) (*excelize.File, error) {
	if spreadsheetData == nil {
		return nil, fmt.Errorf("undefined spreadsheet data")
	}

	if spreadsheetData.SheetStart <= 0 {
		return nil, fmt.Errorf("invalid sheet start index (must be > 0): %d", spreadsheetData.SheetStart)
	}

	if spreadsheetData.Data == nil {
		return nil, fmt.Errorf("undefined results")
	}

	if len(spreadsheetData.Data.Result) == 0 {
		return nil, fmt.Errorf("empty result dataset")
	}

	excelFile := excelize.NewFile()
	if excelFile == nil {
		return nil, fmt.Errorf("excelize.NewFile() failed")
	}

	if spreadsheetMetadata != nil {
		err := addMetadataToSpreadsheet(excelFile, spreadsheetMetadata)
		if err != nil {
			return nil, fmt.Errorf("addMetadataToSpreadsheet() failed: %w", err)
		}
	}

	err := addDataToSpreadsheet(excelFile, spreadsheetData)
	if err != nil {
		return nil, fmt.Errorf("addDataToSpreadsheet() failed: %w", err)
	}

	return excelFile, nil
}

// NewExcelSheetsWithLabels creates a new Excel spreadsheet at the provided excelFilePath that
// will store all the metadata and data that is passed on. The metadata is optional (can be nil)
// but the data must be valid.
func NewExcelSheetsWithLabels(excelFilePath string, spreadsheetMetadata *SpreadsheetMetadata, spreadsheetData *SpreadsheetData) error {
	excelFile, err := ExcelizeWithLabels(spreadsheetMetadata, spreadsheetData)
	if err != nil {
		return fmt.Errorf("results.Excelize() failed: %w", err)
	}

	err = excelFile.SaveAs(excelFilePath)
	if err != nil {
		return err
	}

	return nil
}
