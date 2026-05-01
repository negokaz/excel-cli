package excel

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// FormatRange applies cell styles to the specified range in the worksheet.
// styles is a 2D slice where each element corresponds to a cell in the range.
// A nil element means no style change for that cell.
// The dimensions of styles must match the range exactly.
func FormatRange(workbook Excel, sheetName, rangeStr string, styles [][]*CellStyle) error {
	startCol, startRow, endCol, endRow, err := ParseRange(rangeStr)
	if err != nil {
		return fmt.Errorf("invalid range: %w", err)
	}

	rowCount := endRow - startRow + 1
	colCount := endCol - startCol + 1

	if len(styles) != rowCount {
		return fmt.Errorf("row count mismatch: range %s expects %d rows, but styles has %d rows", rangeStr, rowCount, len(styles))
	}

	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return fmt.Errorf("sheet %q not found", sheetName)
	}
	defer worksheet.Release()

	for rowIdx, styleRow := range styles {
		if len(styleRow) != colCount {
			return fmt.Errorf("column count mismatch: range %s expects %d columns, but row %d has %d columns", rangeStr, colCount, rowIdx, len(styleRow))
		}

		for colIdx, style := range styleRow {
			if style == nil {
				continue
			}
			cell, err := excelize.CoordinatesToCellName(startCol+colIdx, startRow+rowIdx)
			if err != nil {
				return fmt.Errorf("failed to compute cell name: %w", err)
			}
			if err := worksheet.SetCellStyle(cell, style); err != nil {
				return fmt.Errorf("failed to set style for cell %s: %w", cell, err)
			}
		}
	}

	return workbook.Save()
}
