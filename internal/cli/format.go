package cli

import (
	"encoding/json"
	"fmt"
	"path/filepath"
	"strings"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runFormat(args []string) error {
	if len(args) != 4 {
		return fmt.Errorf("usage: excel-cli format <file> <sheet> <range> <styles>")
	}

	filePath := args[0]
	sheetName := args[1]
	rangeStr := args[2]
	stylesJSON := args[3]

	if strings.Contains(rangeStr, "!") {
		return fmt.Errorf("invalid range %q: must not contain sheet name (e.g. 'Sheet1!A1:C3' is not allowed, use the <sheet> argument instead)", rangeStr)
	}

	styles, err := parseStylesJSON(stylesJSON)
	if err != nil {
		return err
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

	return excel.FormatRange(workbook, sheetName, rangeStr, styles)
}

func parseStylesJSON(raw string) ([][]*excel.CellStyle, error) {
	var outer [][]json.RawMessage
	if err := json.Unmarshal([]byte(raw), &outer); err != nil {
		return nil, fmt.Errorf("failed to parse styles as JSON: %w", err)
	}

	result := make([][]*excel.CellStyle, len(outer))
	for i, row := range outer {
		result[i] = make([]*excel.CellStyle, len(row))
		for j, raw := range row {
			if string(raw) == "null" {
				result[i][j] = nil
				continue
			}
			var style excel.CellStyle
			if err := json.Unmarshal(raw, &style); err != nil {
				return nil, fmt.Errorf("failed to parse style at row %d, column %d: %w", i, j, err)
			}
			result[i][j] = &style
		}
	}

	return result, nil
}
