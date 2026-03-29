package cli

import (
	"encoding/json"
	"flag"
	"fmt"
	"path/filepath"
	"strings"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runWrite(args []string) error {
	if len(args) < 4 {
		return fmt.Errorf("usage: excel-cli write <file> <sheet> <range> <values> [--newsheet]")
	}

	filePath := args[0]
	sheetName := args[1]
	rangeStr := args[2]
	valuesJSON := args[3]

	fs := flag.NewFlagSet("write", flag.ContinueOnError)
	newSheet := fs.Bool("newsheet", false, "Create a new sheet before writing")

	if err := fs.Parse(args[4:]); err != nil {
		return err
	}
	if len(fs.Args()) != 0 {
		return fmt.Errorf("usage: excel-cli write <file> <sheet> <range> <values> [--newsheet]")
	}

	if strings.Contains(rangeStr, "!") {
		return fmt.Errorf("invalid range %q: must not contain sheet name (e.g. 'Sheet1!A1:C3' is not allowed, use the <sheet> argument instead)", rangeStr)
	}

	values, err := parseValuesJSON(valuesJSON)
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

	return excel.WriteSheet(workbook, sheetName, rangeStr, values, *newSheet)
}

func parseValuesJSON(raw string) ([][]any, error) {
	var parsed any
	if err := json.Unmarshal([]byte(raw), &parsed); err != nil {
		return nil, fmt.Errorf("failed to parse values as JSON: %w", err)
	}

	outer, ok := parsed.([]any)
	if !ok {
		return nil, fmt.Errorf("values must be a JSON 2-dimensional array")
	}

	result := make([][]any, len(outer))
	for i, item := range outer {
		row, ok := item.([]any)
		if !ok {
			return nil, fmt.Errorf("values must be a JSON 2-dimensional array: row %d is not an array", i)
		}
		result[i] = row
	}

	return result, nil
}
