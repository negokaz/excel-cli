package excel

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
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

	for rowIdx, row := range values {
		for colIdx, val := range row {
			cellCol := startCol + colIdx
			cellRow := startRow + rowIdx
			cellName, err := excelize.CoordinatesToCellName(cellCol, cellRow)
			if err != nil {
				return fmt.Errorf("failed to compute cell name: %w", err)
			}
			if strVal, ok := val.(string); ok && strings.HasPrefix(strVal, "=") {
				if err := worksheet.SetFormula(cellName, strVal); err != nil {
					return fmt.Errorf("failed to set formula at %s: %w", cellName, err)
				}
			} else {
				if err := worksheet.SetValue(cellName, val); err != nil {
					return fmt.Errorf("failed to set value at %s: %w", cellName, err)
				}
			}
		}
	}

	return workbook.Save()
}
