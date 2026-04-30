package cli

import (
	"crypto/md5"
	"fmt"
	"html"
	"slices"
	"strconv"
	"strings"
	"time"

	"github.com/goccy/go-yaml"
	"github.com/xuri/excelize/v2"

	"github.com/negokaz/excel-cli/internal/excel"
)

// HTMLPageParams holds parameters for building a complete HTML page.
type HTMLPageParams struct {
	FilePath    string
	SheetName   string
	UsedRange   string
	Backend     string
	GeneratedAt time.Time
	TableHTML   string
}

func buildHTMLPage(p HTMLPageParams) string {
	var sb strings.Builder
	sb.WriteString(`<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>`)
	sb.WriteString(html.EscapeString(p.SheetName))
	sb.WriteString(` - `)
	sb.WriteString(html.EscapeString(p.FilePath))
	sb.WriteString(`</title>
<style>
body { font-family: sans-serif; margin: 1rem; font-size: 13px; }
h1 { font-size: 1.2rem; margin-bottom: 0.5rem; }
.meta { color: #555; margin-bottom: 1rem; font-size: 0.85rem; }
.meta span { margin-right: 1.5rem; }
table { border-collapse: collapse; white-space: nowrap; }
th, td { border: 1px solid #ccc; padding: 2px 6px; min-width: 40px; }
th { background: #f0f0f0; font-weight: bold; text-align: center; }
tr:hover td { background: #fffbe6; }
</style>
</head>
<body>
<h1>`)
	sb.WriteString(html.EscapeString(p.SheetName))
	sb.WriteString(`</h1>
<div class="meta">
  <span>File: `)
	sb.WriteString(html.EscapeString(p.FilePath))
	sb.WriteString(`</span>
  <span>Range: `)
	sb.WriteString(html.EscapeString(p.UsedRange))
	sb.WriteString(`</span>
  <span>Backend: `)
	sb.WriteString(html.EscapeString(p.Backend))
	sb.WriteString(`</span>
  <span>Generated: `)
	sb.WriteString(html.EscapeString(p.GeneratedAt.UTC().Format(time.RFC3339)))
	sb.WriteString(`</span>
</div>
`)
	sb.WriteString(p.TableHTML)
	sb.WriteString(`
</body>
</html>
`)
	return sb.String()
}

// --- HTML table generation ---

type styleRegistry struct {
	borderStyles   map[string]string
	borderHashToID map[string]string
	borderCounter  int

	fontStyles   map[string]string
	fontHashToID map[string]string
	fontCounter  int

	fillStyles   map[string]string
	fillHashToID map[string]string
	fillCounter  int

	numFmtStyles   map[string]string
	numFmtHashToID map[string]string
	numFmtCounter  int

	decimalStyles   map[string]string
	decimalHashToID map[string]string
	decimalCounter  int
}

func newStyleRegistry() *styleRegistry {
	return &styleRegistry{
		borderStyles:    make(map[string]string),
		borderHashToID:  make(map[string]string),
		fontStyles:      make(map[string]string),
		fontHashToID:    make(map[string]string),
		fillStyles:      make(map[string]string),
		fillHashToID:    make(map[string]string),
		numFmtStyles:    make(map[string]string),
		numFmtHashToID:  make(map[string]string),
		decimalStyles:   make(map[string]string),
		decimalHashToID: make(map[string]string),
	}
}

func (sr *styleRegistry) registerStyle(cellStyle *excel.CellStyle) []string {
	if cellStyle == nil || sr.isEmptyStyle(cellStyle) {
		return []string{}
	}

	var styleIDs []string
	if len(cellStyle.Border) > 0 {
		if borderID := sr.registerBorderStyle(cellStyle.Border); borderID != "" {
			styleIDs = append(styleIDs, borderID)
		}
	}
	if cellStyle.Font != nil {
		if fontID := sr.registerFontStyle(cellStyle.Font); fontID != "" {
			styleIDs = append(styleIDs, fontID)
		}
	}
	if cellStyle.Fill != nil && cellStyle.Fill.Type != "" {
		if fillID := sr.registerFillStyle(cellStyle.Fill); fillID != "" {
			styleIDs = append(styleIDs, fillID)
		}
	}
	if cellStyle.NumFmt != nil && *cellStyle.NumFmt != "" {
		if numFmtID := sr.registerNumFmtStyle(*cellStyle.NumFmt); numFmtID != "" {
			styleIDs = append(styleIDs, numFmtID)
		}
	}
	if cellStyle.DecimalPlaces != nil && *cellStyle.DecimalPlaces != 0 {
		if decimalID := sr.registerDecimalStyle(*cellStyle.DecimalPlaces); decimalID != "" {
			styleIDs = append(styleIDs, decimalID)
		}
	}

	return styleIDs
}

func (sr *styleRegistry) isEmptyStyle(style *excel.CellStyle) bool {
	if len(style.Border) > 0 || style.Font != nil || (style.NumFmt != nil && *style.NumFmt != "") || (style.DecimalPlaces != nil && *style.DecimalPlaces != 0) {
		return false
	}
	return !(style.Fill != nil && style.Fill.Type != "")
}

func (sr *styleRegistry) registerBorderStyle(borders []excel.Border) string {
	if len(borders) == 0 {
		return ""
	}
	yamlStr := convertToYAMLFlow(borders)
	if yamlStr == "" {
		return ""
	}
	styleHash := calculateYAMLHash(yamlStr)
	if styleHash == "" {
		return ""
	}
	if existingID, ok := sr.borderHashToID[styleHash]; ok {
		return existingID
	}
	sr.borderCounter++
	styleID := fmt.Sprintf("b%d", sr.borderCounter)
	sr.borderStyles[styleID] = yamlStr
	sr.borderHashToID[styleHash] = styleID
	return styleID
}

func (sr *styleRegistry) registerFontStyle(font *excel.FontStyle) string {
	if font == nil {
		return ""
	}
	yamlStr := convertToYAMLFlow(font)
	if yamlStr == "" {
		return ""
	}
	styleHash := calculateYAMLHash(yamlStr)
	if styleHash == "" {
		return ""
	}
	if existingID, ok := sr.fontHashToID[styleHash]; ok {
		return existingID
	}
	sr.fontCounter++
	styleID := fmt.Sprintf("f%d", sr.fontCounter)
	sr.fontStyles[styleID] = yamlStr
	sr.fontHashToID[styleHash] = styleID
	return styleID
}

func (sr *styleRegistry) registerFillStyle(fill *excel.FillStyle) string {
	if fill == nil || fill.Type == "" {
		return ""
	}
	yamlStr := convertToYAMLFlow(fill)
	if yamlStr == "" {
		return ""
	}
	styleHash := calculateYAMLHash(yamlStr)
	if styleHash == "" {
		return ""
	}
	if existingID, ok := sr.fillHashToID[styleHash]; ok {
		return existingID
	}
	sr.fillCounter++
	styleID := fmt.Sprintf("l%d", sr.fillCounter)
	sr.fillStyles[styleID] = yamlStr
	sr.fillHashToID[styleHash] = styleID
	return styleID
}

func (sr *styleRegistry) registerNumFmtStyle(numFmt string) string {
	if numFmt == "" {
		return ""
	}
	styleHash := calculateYAMLHash(numFmt)
	if styleHash == "" {
		return ""
	}
	if existingID, ok := sr.numFmtHashToID[styleHash]; ok {
		return existingID
	}
	sr.numFmtCounter++
	styleID := fmt.Sprintf("n%d", sr.numFmtCounter)
	sr.numFmtStyles[styleID] = numFmt
	sr.numFmtHashToID[styleHash] = styleID
	return styleID
}

func (sr *styleRegistry) registerDecimalStyle(decimal int) string {
	if decimal == 0 {
		return ""
	}
	yamlStr := convertToYAMLFlow(decimal)
	if yamlStr == "" {
		return ""
	}
	styleHash := calculateYAMLHash(yamlStr)
	if styleHash == "" {
		return ""
	}
	if existingID, ok := sr.decimalHashToID[styleHash]; ok {
		return existingID
	}
	sr.decimalCounter++
	styleID := fmt.Sprintf("d%d", sr.decimalCounter)
	sr.decimalStyles[styleID] = yamlStr
	sr.decimalHashToID[styleHash] = styleID
	return styleID
}

func (sr *styleRegistry) generateStyleDefinitions() string {
	totalCount := len(sr.borderStyles) + len(sr.fontStyles) + len(sr.fillStyles) + len(sr.numFmtStyles) + len(sr.decimalStyles)
	if totalCount == 0 {
		return ""
	}

	var result strings.Builder
	result.WriteString("<h2>Style Definitions</h2>\n")
	result.WriteString("<div class=\"style-definitions\">\n")
	result.WriteString(sr.generateStyleDefTag(sr.borderStyles, "border"))
	result.WriteString(sr.generateStyleDefTag(sr.fontStyles, "font"))
	result.WriteString(sr.generateStyleDefTag(sr.fillStyles, "fill"))
	result.WriteString(sr.generateStyleDefTag(sr.numFmtStyles, "numFmt"))
	result.WriteString(sr.generateStyleDefTag(sr.decimalStyles, "decimalPlaces"))
	result.WriteString("</div>\n\n")
	return result.String()
}

func (sr *styleRegistry) generateStyleDefTag(styles map[string]string, styleLabel string) string {
	if len(styles) == 0 {
		return ""
	}

	styleIDs := make([]string, 0, len(styles))
	for styleID := range styles {
		styleIDs = append(styleIDs, styleID)
	}
	sortStyleIDs(styleIDs)

	var result strings.Builder
	for _, styleID := range styleIDs {
		yamlStr := styles[styleID]
		if yamlStr != "" {
			result.WriteString(fmt.Sprintf("<code class=\"style language-yaml\" id=\"%s\">%s: %s</code>\n", styleID, styleLabel, html.EscapeString(yamlStr)))
		}
	}
	return result.String()
}

func sortStyleIDs(styleIDs []string) {
	slices.SortFunc(styleIDs, func(a, b string) int {
		ai, _ := strconv.Atoi(a[1:])
		bi, _ := strconv.Atoi(b[1:])
		return ai - bi
	})
}

func calculateYAMLHash(yamlStr string) string {
	if yamlStr == "" {
		return ""
	}
	hash := md5.Sum([]byte(yamlStr))
	return fmt.Sprintf("%x", hash)[:8]
}

func convertToYAMLFlow(value any) string {
	if value == nil {
		return ""
	}
	yamlBytes, err := yaml.MarshalWithOptions(value, yaml.Flow(true), yaml.OmitEmpty())
	if err != nil {
		return ""
	}
	return strings.TrimSpace(strings.ReplaceAll(string(yamlBytes), "\"", ""))
}

func createHTMLTableOfValues(ws excel.Worksheet, startCol, startRow, endCol, endRow int) (string, error) {
	return createHTMLTable(startCol, startRow, endCol, endRow,
		func(cell string) (string, error) { return ws.GetValue(cell) },
		ws.GetMergedCells,
	)
}

func createHTMLTableOfFormula(ws excel.Worksheet, startCol, startRow, endCol, endRow int) (string, error) {
	return createHTMLTable(startCol, startRow, endCol, endRow,
		func(cell string) (string, error) { return ws.GetFormula(cell) },
		ws.GetMergedCells,
	)
}

func createHTMLTableOfValuesWithStyle(ws excel.Worksheet, startCol, startRow, endCol, endRow int) (string, error) {
	return createHTMLTableWithStyle(startCol, startRow, endCol, endRow,
		func(cell string) (string, error) { return ws.GetValue(cell) },
		func(cell string) (*excel.CellStyle, error) { return ws.GetCellStyle(cell) },
		ws.GetMergedCells,
	)
}

func createHTMLTableOfFormulaWithStyle(ws excel.Worksheet, startCol, startRow, endCol, endRow int) (string, error) {
	return createHTMLTableWithStyle(startCol, startRow, endCol, endRow,
		func(cell string) (string, error) { return ws.GetFormula(cell) },
		func(cell string) (*excel.CellStyle, error) { return ws.GetCellStyle(cell) },
		ws.GetMergedCells,
	)
}

func createHTMLTable(
	startCol, startRow, endCol, endRow int,
	extractor func(string) (string, error),
	mergedCellsGetter func() ([]excel.MergedCell, error),
) (string, error) {
	return createHTMLTableWithStyle(startCol, startRow, endCol, endRow, extractor, nil, mergedCellsGetter)
}

func createHTMLTableWithStyle(
	startCol, startRow, endCol, endRow int,
	extractor func(string) (string, error),
	styleExtractor func(string) (*excel.CellStyle, error),
	mergedCellsGetter func() ([]excel.MergedCell, error),
) (string, error) {
	registry := newStyleRegistry()

	type cellKey struct{ col, row int }
	type mergeSpan struct{ colspan, rowspan int }

	mergeSpanMap := make(map[cellKey]mergeSpan)
	skipCells := make(map[cellKey]struct{})

	if mergedCellsGetter != nil {
		mergedCells, err := mergedCellsGetter()
		if err != nil {
			return "", fmt.Errorf("failed to get merged cells: %w", err)
		}
		for _, mc := range mergedCells {
			colspan := mc.EndCol - mc.StartCol + 1
			rowspan := mc.EndRow - mc.StartRow + 1
			if colspan <= 1 && rowspan <= 1 {
				continue
			}
			mergeSpanMap[cellKey{mc.StartCol, mc.StartRow}] = mergeSpan{colspan, rowspan}
			for r := mc.StartRow; r <= mc.EndRow; r++ {
				for c := mc.StartCol; c <= mc.EndCol; c++ {
					if r == mc.StartRow && c == mc.StartCol {
						continue
					}
					skipCells[cellKey{c, r}] = struct{}{}
				}
			}
		}
	}

	var table strings.Builder
	table.WriteString("<table>\n<tr><th></th>")
	for col := startCol; col <= endCol; col++ {
		name, _ := excelize.ColumnNumberToName(col)
		table.WriteString(fmt.Sprintf("<th>%s</th>", name))
	}
	table.WriteString("</tr>\n")

	for row := startRow; row <= endRow; row++ {
		table.WriteString("<tr>")
		table.WriteString(fmt.Sprintf("<th>%d</th>", row))
		for col := startCol; col <= endCol; col++ {
			if _, skip := skipCells[cellKey{col, row}]; skip {
				continue
			}
			axis, _ := excelize.CoordinatesToCellName(col, row)
			value, _ := extractor(axis)

			tdTag := "<td"
			if span, ok := mergeSpanMap[cellKey{col, row}]; ok {
				if span.colspan > 1 {
					tdTag += fmt.Sprintf(` colspan="%d"`, span.colspan)
				}
				if span.rowspan > 1 {
					tdTag += fmt.Sprintf(` rowspan="%d"`, span.rowspan)
				}
			}
			if styleExtractor != nil {
				cellStyle, err := styleExtractor(axis)
				if err == nil && cellStyle != nil {
					styleIDs := registry.registerStyle(cellStyle)
					if len(styleIDs) > 0 {
						tdTag += fmt.Sprintf(` style-ref="%s"`, strings.Join(styleIDs, " "))
					}
				}
			}
			tdTag += ">"
			table.WriteString(fmt.Sprintf("%s%s</td>", tdTag, strings.ReplaceAll(html.EscapeString(value), "\n", "<br>")))
		}
		table.WriteString("</tr>\n")
	}
	table.WriteString("</table>")

	var finalResult strings.Builder
	styleDefinitions := registry.generateStyleDefinitions()
	if styleDefinitions != "" {
		finalResult.WriteString(styleDefinitions)
	}
	finalResult.WriteString("<h2>Sheet Data</h2>\n")
	finalResult.WriteString(table.String())
	return finalResult.String(), nil
}
