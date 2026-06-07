package cli

import (
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runWrite(args []string) error {
	if len(args) < 2 {
		return fmt.Errorf("usage: excel-cli write <file> <path> (--value <json|-> | --formula <json|-> | --style <json|-> | --props <json|->)")
	}

	fs := flag.NewFlagSet("write", flag.ContinueOnError)
	valuePayload := fs.String("value", "", "Write normal values")
	formulaPayload := fs.String("formula", "", "Write formulas or normal values")
	stylePayload := fs.String("style", "", "Apply styles")
	propsPayload := fs.String("props", "", "Update properties")

	filePath := args[0]
	rawPath := args[1]
	if err := fs.Parse(args[2:]); err != nil {
		return err
	}
	if len(fs.Args()) != 0 {
		return fmt.Errorf("usage: excel-cli write <file> <path> (--value <json|-> | --formula <json|-> | --style <json|-> | --props <json|->)")
	}

	target, err := parseTargetPath(rawPath)
	if err != nil {
		return err
	}

	channels := []struct {
		name    string
		payload string
	}{
		{name: "value", payload: *valuePayload},
		{name: "formula", payload: *formulaPayload},
		{name: "style", payload: *stylePayload},
		{name: "props", payload: *propsPayload},
	}
	selected := ""
	for _, channel := range channels {
		if channel.payload == "" {
			continue
		}
		if selected != "" {
			return fmt.Errorf("zero or multiple update channels are provided")
		}
		selected = channel.name
	}
	if selected == "" {
		return fmt.Errorf("zero or multiple update channels are provided")
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

	switch selected {
	case "value":
		if target.Kind != pathKindRange {
			return fmt.Errorf("unsupported path kind for channel value")
		}
		payload, err := resolveWriteJSONPayload(*valuePayload)
		if err != nil {
			return err
		}
		values, err := parseValuesJSON(payload)
		if err != nil {
			return err
		}
		if err := validateMatrixShape(target.RangeRef, values); err != nil {
			return err
		}
		worksheet, err := workbook.FindSheet(target.SheetName)
		if err != nil {
			return fmt.Errorf("sheet not found: %s", target.SheetName)
		}
		defer worksheet.Release()
		if err := worksheet.SetValuesRange(target.RangeRef, values); err != nil {
			return fmt.Errorf("failed to write values: %w", err)
		}
	case "formula":
		if target.Kind != pathKindRange {
			return fmt.Errorf("unsupported path kind for channel formula")
		}
		payload, err := resolveWriteJSONPayload(*formulaPayload)
		if err != nil {
			return err
		}
		values, err := parseValuesJSON(payload)
		if err != nil {
			return err
		}
		if err := validateMatrixShape(target.RangeRef, values); err != nil {
			return err
		}
		worksheet, err := workbook.FindSheet(target.SheetName)
		if err != nil {
			return fmt.Errorf("sheet not found: %s", target.SheetName)
		}
		defer worksheet.Release()
		if err := worksheet.SetFormulasRange(target.RangeRef, values); err != nil {
			return fmt.Errorf("failed to write formulas: %w", err)
		}
	case "style":
		if target.Kind != pathKindRange {
			return fmt.Errorf("unsupported path kind for channel style")
		}
		payload, err := resolveWriteJSONPayload(*stylePayload)
		if err != nil {
			return err
		}
		styles, err := parseStylesJSON(payload)
		if err != nil {
			return err
		}
		if err := validateStyles(styles); err != nil {
			return err
		}
		if err := validateStyleMatrixShape(target.RangeRef, styles); err != nil {
			return err
		}
		if err := excel.FormatRange(workbook, target.SheetName, target.RangeRef, styles); err != nil {
			return err
		}
		return writeJSON(struct {
			Path    string `json:"path"`
			Kind    string `json:"kind"`
			Action  string `json:"action"`
			Channel string `json:"channel"`
		}{
			Path:    target.Canonical(),
			Kind:    "range",
			Action:  "write",
			Channel: "style",
		})
	case "props":
		if target.Kind != pathKindSheet {
			return fmt.Errorf("unsupported path kind for channel props")
		}
		payload, err := resolveWriteJSONPayload(*propsPayload)
		if err != nil {
			return err
		}
		props, err := parseSheetPropsJSON(payload)
		if err != nil {
			return err
		}
		worksheet, err := workbook.FindSheet(target.SheetName)
		if err != nil {
			return fmt.Errorf("sheet not found: %s", target.SheetName)
		}
		defer worksheet.Release()
		if props.Hidden != nil {
			if err := worksheet.SetHidden(*props.Hidden); err != nil {
				return fmt.Errorf("failed to update sheet properties: %w", err)
			}
		}
	}

	if err := workbook.Save(); err != nil {
		return fmt.Errorf("failed to save workbook: %w", err)
	}

	kind := "range"
	if target.Kind == pathKindSheet {
		kind = "sheet"
	}
	return writeJSON(struct {
		Path    string `json:"path"`
		Kind    string `json:"kind"`
		Action  string `json:"action"`
		Channel string `json:"channel"`
	}{
		Path:    target.Canonical(),
		Kind:    kind,
		Action:  "write",
		Channel: selected,
	})
}

func resolveWriteJSONPayload(raw string) (string, error) {
	if raw != "-" {
		return raw, nil
	}

	data, err := io.ReadAll(os.Stdin)
	if err != nil {
		return "", fmt.Errorf("failed to read JSON from stdin: %w", err)
	}
	return string(data), nil
}

type sheetProps struct {
	Hidden *bool
}

func parseSheetPropsJSON(raw string) (sheetProps, error) {
	var payload map[string]json.RawMessage
	if err := json.Unmarshal([]byte(raw), &payload); err != nil {
		return sheetProps{}, fmt.Errorf("failed to parse props as JSON: %w", err)
	}
	var result sheetProps
	for key, value := range payload {
		switch key {
		case "hidden":
			var hidden bool
			if err := json.Unmarshal(value, &hidden); err != nil {
				return sheetProps{}, fmt.Errorf("hidden must be a boolean")
			}
			result.Hidden = &hidden
		default:
			return sheetProps{}, fmt.Errorf("unsupported property: %s", key)
		}
	}
	return result, nil
}

func validateMatrixShape(rangeRef string, values [][]any) error {
	startCol, startRow, endCol, endRow, err := excel.ParseRange(rangeRef)
	if err != nil {
		return err
	}
	expectedRows := endRow - startRow + 1
	expectedCols := endCol - startCol + 1
	if len(values) != expectedRows {
		return fmt.Errorf("row count mismatch: range %s expects %d rows, but values has %d rows", rangeRef, expectedRows, len(values))
	}
	for rowIdx, row := range values {
		if len(row) != expectedCols {
			return fmt.Errorf("column count mismatch: range %s expects %d columns, but row %d has %d columns", rangeRef, expectedCols, rowIdx, len(row))
		}
	}
	return nil
}

func validateStyleMatrixShape(rangeRef string, styles [][]*excel.CellStyle) error {
	startCol, startRow, endCol, endRow, err := excel.ParseRange(rangeRef)
	if err != nil {
		return err
	}
	expectedRows := endRow - startRow + 1
	expectedCols := endCol - startCol + 1
	if len(styles) != expectedRows {
		return fmt.Errorf("row count mismatch: range %s expects %d rows, but styles has %d rows", rangeRef, expectedRows, len(styles))
	}
	for rowIdx, row := range styles {
		if len(row) != expectedCols {
			return fmt.Errorf("column count mismatch: range %s expects %d columns, but row %d has %d columns", rangeRef, expectedCols, rowIdx, len(row))
		}
	}
	return nil
}
