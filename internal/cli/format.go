package cli

import (
	"encoding/json"
	"fmt"
	"path/filepath"
	"regexp"
	"slices"
	"strings"

	"github.com/negokaz/excel-cli/internal/excel"
)

var excelStyleColorPattern = regexp.MustCompile(`^#[0-9A-Fa-f]{6}$`)

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
	if err := validateStyles(styles); err != nil {
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

func validateStyles(styles [][]*excel.CellStyle) error {
	for rowIdx, row := range styles {
		for colIdx, style := range row {
			if style == nil {
				continue
			}
			if err := validateCellStyle(style, rowIdx, colIdx); err != nil {
				return err
			}
		}
	}
	return nil
}

func validateCellStyle(style *excel.CellStyle, rowIdx, colIdx int) error {
	location := fmt.Sprintf("row %d, column %d", rowIdx, colIdx)

	for borderIdx, border := range style.Border {
		if !containsEnum(excel.BorderTypeValues(), border.Type) {
			return fmt.Errorf("invalid border.type at %s, border %d: %q", location, borderIdx, border.Type)
		}
		if border.Style != "" && !containsEnum(excel.BorderStyleValues(), border.Style) {
			return fmt.Errorf("invalid border.style at %s, border %d: %q", location, borderIdx, border.Style)
		}
		if border.Color != "" && !excelStyleColorPattern.MatchString(border.Color) {
			return fmt.Errorf("invalid border.color at %s, border %d: %q", location, borderIdx, border.Color)
		}
	}

	if style.Font != nil {
		if style.Font.Underline != nil && !containsEnum(excel.FontUnderlineValues(), *style.Font.Underline) {
			return fmt.Errorf("invalid font.underline at %s: %q", location, *style.Font.Underline)
		}
		if style.Font.Size != nil && (*style.Font.Size < 1 || *style.Font.Size > 409) {
			return fmt.Errorf("invalid font.size at %s: %d", location, *style.Font.Size)
		}
		if style.Font.Color != nil && !excelStyleColorPattern.MatchString(*style.Font.Color) {
			return fmt.Errorf("invalid font.color at %s: %q", location, *style.Font.Color)
		}
		if style.Font.VertAlign != nil && !containsEnum(excel.FontVertAlignValues(), *style.Font.VertAlign) {
			return fmt.Errorf("invalid font.vertAlign at %s: %q", location, *style.Font.VertAlign)
		}
	}

	if style.Fill != nil {
		if style.Fill.Type != "" && !containsEnum(excel.FillTypeValues(), style.Fill.Type) {
			return fmt.Errorf("invalid fill.type at %s: %q", location, style.Fill.Type)
		}
		if style.Fill.Pattern != "" && !containsEnum(excel.FillPatternValues(), style.Fill.Pattern) {
			return fmt.Errorf("invalid fill.pattern at %s: %q", location, style.Fill.Pattern)
		}
		for colorIdx, color := range style.Fill.Color {
			if !excelStyleColorPattern.MatchString(color) {
				return fmt.Errorf("invalid fill.color at %s, color %d: %q", location, colorIdx, color)
			}
		}
		if style.Fill.Shading != nil && !containsEnum(excel.FillShadingValues(), *style.Fill.Shading) {
			return fmt.Errorf("invalid fill.shading at %s: %q", location, *style.Fill.Shading)
		}
	}

	if style.DecimalPlaces != nil && (*style.DecimalPlaces < 0 || *style.DecimalPlaces > 30) {
		return fmt.Errorf("invalid decimalPlaces at %s: %d", location, *style.DecimalPlaces)
	}

	return nil
}

func containsEnum[T ~string](values []T, want T) bool {
	return slices.Contains(values, want)
}
