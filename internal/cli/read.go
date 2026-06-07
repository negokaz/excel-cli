package cli

import (
	"flag"
	"fmt"
	"path/filepath"

	"github.com/xuri/excelize/v2"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runRead(args []string) error {
	if len(args) < 2 {
		return fmt.Errorf("usage: excel-cli read <file> <path> [--value | --formula | --style]")
	}

	fs := flag.NewFlagSet("read", flag.ContinueOnError)
	showValue := fs.Bool("value", false, "Return displayed values for a range")
	showFormula := fs.Bool("formula", false, "Return formulas for a range")
	showStyle := fs.Bool("style", false, "Return styles for a range")

	filePath := args[0]
	rawPath := args[1]
	if err := fs.Parse(args[2:]); err != nil {
		return err
	}
	if len(fs.Args()) != 0 {
		return fmt.Errorf("usage: excel-cli read <file> <path> [--value | --formula | --style]")
	}

	target, err := parseTargetPath(rawPath)
	if err != nil {
		return err
	}

	selected := 0
	if *showValue {
		selected++
	}
	if *showFormula {
		selected++
	}
	if *showStyle {
		selected++
	}
	if selected > 1 {
		return fmt.Errorf("more than one of --value, --formula, or --style was provided")
	}

	absPath, err := filepath.Abs(filePath)
	if err != nil {
		return fmt.Errorf("failed to resolve path: %w", err)
	}

	workbook, release, err := excel.OpenFile(absPath)
	if err != nil {
		return fmt.Errorf("failed to open file: %w", err)
	}
	defer release()

	switch target.Kind {
	case pathKindWorkbook:
		sheets, err := workbook.GetSheets()
		if err != nil {
			return fmt.Errorf("failed to get sheets: %w", err)
		}
		for _, sheet := range sheets {
			defer sheet.Release()
		}
		return writeJSON(struct {
			Path    string `json:"path"`
			Kind    string `json:"kind"`
			Backend string `json:"backend"`
			Data    any    `json:"data"`
		}{
			Path:    target.Canonical(),
			Kind:    "workbook",
			Backend: workbook.GetBackendName(),
			Data: struct {
				SheetCount int `json:"sheetCount"`
			}{SheetCount: len(sheets)},
		})
	case pathKindSheet:
		worksheet, err := workbook.FindSheet(target.SheetName)
		if err != nil {
			return fmt.Errorf("sheet not found: %s", target.SheetName)
		}
		defer worksheet.Release()
		hidden, err := worksheet.IsHidden()
		if err != nil {
			return fmt.Errorf("failed to get sheet metadata: %w", err)
		}
		usedRange, err := worksheet.GetDimention()
		if err != nil {
			return fmt.Errorf("failed to get sheet metadata: %w", err)
		}
		name, err := worksheet.Name()
		if err != nil {
			return fmt.Errorf("failed to get sheet metadata: %w", err)
		}
		return writeJSON(struct {
			Path    string `json:"path"`
			Kind    string `json:"kind"`
			Backend string `json:"backend"`
			Data    any    `json:"data"`
		}{
			Path:    target.Canonical(),
			Kind:    "sheet",
			Backend: workbook.GetBackendName(),
			Data: struct {
				Name      string `json:"name"`
				Hidden    bool   `json:"hidden"`
				UsedRange string `json:"usedRange"`
			}{
				Name:      name,
				Hidden:    hidden,
				UsedRange: excel.NormalizeRange(usedRange),
			},
		})
	case pathKindRange:
		worksheet, err := workbook.FindSheet(target.SheetName)
		if err != nil {
			return fmt.Errorf("sheet not found: %s", target.SheetName)
		}
		defer worksheet.Release()

		channel := "value"
		if *showFormula {
			channel = "formula"
		}
		if *showStyle {
			channel = "style"
		}

		switch channel {
		case "value":
			values, err := worksheet.GetValuesRangeAny(target.RangeRef)
			if err != nil {
				return fmt.Errorf("failed to read range: %w", err)
			}
			return writeJSON(struct {
				Path    string `json:"path"`
				Kind    string `json:"kind"`
				Backend string `json:"backend"`
				Data    any    `json:"data"`
			}{
				Path:    target.Canonical(),
				Kind:    "range",
				Backend: workbook.GetBackendName(),
				Data: struct {
					Values [][]any `json:"values"`
				}{Values: values},
			})
		case "formula":
			formulas, err := worksheet.GetFormulasRangeAny(target.RangeRef)
			if err != nil {
				return fmt.Errorf("failed to read range: %w", err)
			}
			return writeJSON(struct {
				Path    string `json:"path"`
				Kind    string `json:"kind"`
				Backend string `json:"backend"`
				Data    any    `json:"data"`
			}{
				Path:    target.Canonical(),
				Kind:    "range",
				Backend: workbook.GetBackendName(),
				Data: struct {
					Formulas [][]any `json:"formulas"`
				}{Formulas: formulas},
			})
		case "style":
			styles, err := getStyleMatrix(worksheet, target.RangeRef)
			if err != nil {
				return fmt.Errorf("failed to read range style: %w", err)
			}
			return writeJSON(struct {
				Path    string `json:"path"`
				Kind    string `json:"kind"`
				Backend string `json:"backend"`
				Data    any    `json:"data"`
			}{
				Path:    target.Canonical(),
				Kind:    "range",
				Backend: workbook.GetBackendName(),
				Data: struct {
					Styles [][]*excel.CellStyle `json:"styles"`
				}{Styles: styles},
			})
		}
	}

	return fmt.Errorf("unsupported path kind")
}

func getStyleMatrix(worksheet excel.Worksheet, rangeRef string) ([][]*excel.CellStyle, error) {
	startCol, startRow, endCol, endRow, err := excel.ParseRange(rangeRef)
	if err != nil {
		return nil, err
	}
	rows := endRow - startRow + 1
	cols := endCol - startCol + 1
	result := make([][]*excel.CellStyle, rows)
	for row := 0; row < rows; row++ {
		result[row] = make([]*excel.CellStyle, cols)
		for col := 0; col < cols; col++ {
			cell, err := excelize.CoordinatesToCellName(startCol+col, startRow+row)
			if err != nil {
				return nil, err
			}
			style, err := worksheet.GetCellStyle(cell)
			if err != nil {
				return nil, err
			}
			if style != nil && isEmptyCellStyle(style) {
				style = nil
			}
			result[row][col] = style
		}
	}
	return result, nil
}

func isEmptyCellStyle(style *excel.CellStyle) bool {
	if style == nil {
		return true
	}
	if len(style.Border) > 0 || style.Font != nil || style.Fill != nil {
		return false
	}
	if style.NumFmt != nil && *style.NumFmt != "" {
		return false
	}
	return style.DecimalPlaces == nil || *style.DecimalPlaces == 0
}
