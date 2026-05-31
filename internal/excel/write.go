package excel

import (
	"fmt"
)

func WriteSheet(workbook Excel, sheetName, rangeStr string, values [][]any, newSheet bool) error {
	startCol, startRow, endCol, endRow, err := ParseRange(rangeStr)
	if err != nil {
		return fmt.Errorf("invalid range: %w", err)
	}

	rowCount := endRow - startRow + 1
	colCount := endCol - startCol + 1
	if len(values) != rowCount {
		return fmt.Errorf("row count mismatch: range %s expects %d rows, but values has %d rows", rangeStr, rowCount, len(values))
	}
	for rowIdx, row := range values {
		if len(row) != colCount {
			return fmt.Errorf("column count mismatch: range %s expects %d columns, but row %d has %d columns", rangeStr, colCount, rowIdx, len(row))
		}
	}

	if newSheet {
		existing, findErr := workbook.FindSheet(sheetName)
		if findErr == nil {
			existing.Release()
			return fmt.Errorf("sheet %q already exists", sheetName)
		}
		if err := workbook.CreateNewSheet(sheetName); err != nil {
			return fmt.Errorf("failed to create sheet %q: %w", sheetName, err)
		}
	}

	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return fmt.Errorf("sheet %q not found", sheetName)
	}
	defer worksheet.Release()

	if err := worksheet.SetValuesRange(rangeStr, values); err != nil {
		return fmt.Errorf("failed to write values: %w", err)
	}

	return workbook.Save()
}
