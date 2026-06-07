package excel

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

type ExcelizeExcel struct {
	file *excelize.File
}

func NewExcelizeExcel(file *excelize.File) Excel {
	return &ExcelizeExcel{file: file}
}

func (e *ExcelizeExcel) GetBackendName() string { return "excelize" }

func (e *ExcelizeExcel) FindSheet(sheetName string) (Worksheet, error) {
	index, err := e.file.GetSheetIndex(sheetName)
	if err != nil {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	if index < 0 {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	return &ExcelizeWorksheet{file: e.file, sheetName: sheetName}, nil
}

func (e *ExcelizeExcel) CreateNewSheet(sheetName string) error {
	_, err := e.file.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to create new sheet: %w", err)
	}
	return nil
}

func (e *ExcelizeExcel) DeleteSheet(sheetName string) error {
	if err := e.file.DeleteSheet(sheetName); err != nil {
		return fmt.Errorf("failed to delete sheet: %w", err)
	}
	return nil
}

func (e *ExcelizeExcel) CopySheet(srcSheetName string, destSheetName string) error {
	srcIndex, err := e.file.GetSheetIndex(srcSheetName)
	if srcIndex < 0 {
		return fmt.Errorf("source sheet not found: %s", srcSheetName)
	}
	if err != nil {
		return err
	}
	destIndex, err := e.file.NewSheet(destSheetName)
	if err != nil {
		return fmt.Errorf("failed to create destination sheet: %w", err)
	}
	if err := e.file.CopySheet(srcIndex, destIndex); err != nil {
		return fmt.Errorf("failed to copy sheet: %w", err)
	}
	srcNext := e.file.GetSheetList()[srcIndex+1]
	if srcNext != srcSheetName {
		e.file.MoveSheet(destSheetName, srcNext)
	}
	return nil
}

func (e *ExcelizeExcel) GetSheets() ([]Worksheet, error) {
	sheetList := e.file.GetSheetList()
	worksheets := make([]Worksheet, len(sheetList))
	for i, sheetName := range sheetList {
		worksheets[i] = &ExcelizeWorksheet{file: e.file, sheetName: sheetName}
	}
	return worksheets, nil
}

// Save saves the Excel file.
// Excelize's Save method restricts the file path length to 207 characters,
// so we bypass it by writing directly to the file.
func (w *ExcelizeExcel) Save() error {
	file, err := os.OpenFile(filepath.Clean(w.file.Path), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if err != nil {
		return err
	}
	defer file.Close()
	return w.file.Write(file)
}

type ExcelizeWorksheet struct {
	file      *excelize.File
	sheetName string
}

func (w *ExcelizeWorksheet) Release() {}

func (w *ExcelizeWorksheet) Name() (string, error) { return w.sheetName, nil }

func (w *ExcelizeWorksheet) IsHidden() (bool, error) {
	visible, err := w.file.GetSheetVisible(w.sheetName)
	if err != nil {
		return false, fmt.Errorf("failed to get sheet visibility: %w", err)
	}
	return !visible, nil
}

func (w *ExcelizeWorksheet) SetHidden(hidden bool) error {
	if err := w.file.SetSheetVisible(w.sheetName, !hidden); err != nil {
		return fmt.Errorf("failed to set sheet visibility: %w", err)
	}
	return nil
}

func (w *ExcelizeWorksheet) GetTables() ([]Table, error) {
	tables, err := w.file.GetTables(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get tables: %w", err)
	}
	tableList := make([]Table, len(tables))
	for i, table := range tables {
		tableList[i] = Table{Name: table.Name, Range: NormalizeRange(table.Range)}
	}
	return tableList, nil
}

func (w *ExcelizeWorksheet) GetPivotTables() ([]PivotTable, error) {
	pivotTables, err := w.file.GetPivotTables(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get pivot tables: %w", err)
	}
	pivotTableList := make([]PivotTable, len(pivotTables))
	for i, pt := range pivotTables {
		pivotTableList[i] = PivotTable{Name: pt.Name, Range: NormalizeRange(pt.PivotTableRange)}
	}
	return pivotTableList, nil
}

func (w *ExcelizeWorksheet) SetValue(cell string, value any) error {
	if err := w.file.SetCellValue(w.sheetName, cell, value); err != nil {
		return err
	}
	return w.updateDimension(cell)
}

func (w *ExcelizeWorksheet) SetFormula(cell string, formula string) error {
	if err := w.file.SetCellFormula(w.sheetName, cell, formula); err != nil {
		return err
	}
	return w.updateDimension(cell)
}

func (w *ExcelizeWorksheet) GetValue(cell string) (string, error) {
	value, err := w.file.GetCellValue(w.sheetName, cell)
	if err != nil {
		return "", err
	}
	if value == "" {
		formula, err := w.file.GetCellFormula(w.sheetName, cell)
		if err != nil {
			return "", fmt.Errorf("failed to get formula: %w", err)
		}
		if formula != "" {
			return w.file.CalcCellValue(w.sheetName, cell)
		}
	}
	return value, nil
}

func (w *ExcelizeWorksheet) GetFormula(cell string) (string, error) {
	formula, err := w.file.GetCellFormula(w.sheetName, cell)
	if err != nil {
		return "", fmt.Errorf("failed to get formula: %w", err)
	}
	if formula == "" {
		return w.GetValue(cell)
	}
	if !strings.HasPrefix(formula, "=") {
		formula = "=" + formula
	}
	return formula, nil
}

func (w *ExcelizeWorksheet) GetDimention() (string, error) {
	return w.file.GetSheetDimension(w.sheetName)
}

func (w *ExcelizeWorksheet) CapturePicture(captureRange string) (string, error) {
	return "", fmt.Errorf("CapturePicture is not supported in Excelize")
}

func (w *ExcelizeWorksheet) AddTable(tableRange, tableName string) error {
	enable := true
	return w.file.AddTable(w.sheetName, &excelize.Table{
		Range:             tableRange,
		Name:              tableName,
		StyleName:         "TableStyleMedium2",
		ShowColumnStripes: true,
		ShowHeaderRow:     &enable,
		ShowRowStripes:    &enable,
	})
}

func (w *ExcelizeWorksheet) GetCellStyle(cell string) (*CellStyle, error) {
	styleID, err := w.file.GetCellStyle(w.sheetName, cell)
	if err != nil {
		return nil, fmt.Errorf("failed to get cell style: %w", err)
	}
	style, err := w.file.GetStyle(styleID)
	if err != nil {
		return nil, fmt.Errorf("failed to get style details: %w", err)
	}
	return convertExcelizeStyleToCellStyle(style), nil
}

func (w *ExcelizeWorksheet) SetCellStyle(cell string, style *CellStyle) error {
	excelizeStyle := convertCellStyleToExcelizeStyle(style)
	styleID, err := w.file.NewStyle(excelizeStyle)
	if err != nil {
		return fmt.Errorf("failed to create style: %w", err)
	}
	if err := w.file.SetCellStyle(w.sheetName, cell, cell, styleID); err != nil {
		return fmt.Errorf("failed to set cell style: %w", err)
	}
	return nil
}

func convertCellStyleToExcelizeStyle(style *CellStyle) *excelize.Style {
	result := &excelize.Style{}
	if len(style.Border) > 0 {
		borders := make([]excelize.Border, len(style.Border))
		for i, border := range style.Border {
			b := excelize.Border{Type: border.Type.String()}
			if border.Color != "" {
				b.Color = strings.TrimPrefix(border.Color, "#")
			}
			b.Style = borderStyleNameToInt(border.Style)
			borders[i] = b
		}
		result.Border = borders
	}
	if style.Font != nil {
		font := &excelize.Font{}
		if style.Font.Bold != nil {
			font.Bold = *style.Font.Bold
		}
		if style.Font.Italic != nil {
			font.Italic = *style.Font.Italic
		}
		if style.Font.Underline != nil {
			font.Underline = style.Font.Underline.String()
		}
		if style.Font.Size != nil && *style.Font.Size > 0 {
			font.Size = float64(*style.Font.Size)
		}
		if style.Font.Strike != nil {
			font.Strike = *style.Font.Strike
		}
		if style.Font.Color != nil && *style.Font.Color != "" {
			font.Color = strings.TrimPrefix(*style.Font.Color, "#")
		}
		if style.Font.VertAlign != nil {
			font.VertAlign = style.Font.VertAlign.String()
		}
		result.Font = font
	}
	if style.Fill != nil {
		fill := excelize.Fill{}
		if style.Fill.Type != "" {
			fill.Type = style.Fill.Type.String()
		}
		fill.Pattern = fillPatternNameToInt(style.Fill.Pattern)
		if len(style.Fill.Color) > 0 {
			colors := make([]string, len(style.Fill.Color))
			for i, color := range style.Fill.Color {
				colors[i] = strings.TrimPrefix(color, "#")
			}
			fill.Color = colors
		}
		if style.Fill.Shading != nil {
			fill.Shading = fillShadingNameToInt(*style.Fill.Shading)
		}
		result.Fill = fill
	}
	if style.NumFmt != nil && *style.NumFmt != "" {
		result.CustomNumFmt = style.NumFmt
	}
	if style.DecimalPlaces != nil && *style.DecimalPlaces > 0 {
		result.DecimalPlaces = style.DecimalPlaces
	}
	return result
}

func convertExcelizeStyleToCellStyle(style *excelize.Style) *CellStyle {
	result := &CellStyle{}
	if len(style.Border) > 0 {
		var borders []Border
		for _, border := range style.Border {
			b := Border{Type: BorderType(border.Type)}
			if border.Color != "" {
				b.Color = "#" + strings.ToUpper(border.Color)
			}
			if border.Style != 0 {
				b.Style = intToBorderStyleName(border.Style)
			}
			borders = append(borders, b)
		}
		result.Border = borders
	}
	if style.Font != nil {
		font := &FontStyle{}
		if style.Font.Bold {
			font.Bold = &style.Font.Bold
		}
		if style.Font.Italic {
			font.Italic = &style.Font.Italic
		}
		if style.Font.Underline != "" {
			u := FontUnderline(style.Font.Underline)
			font.Underline = &u
		}
		if style.Font.Size > 0 {
			size := int(style.Font.Size)
			font.Size = &size
		}
		if style.Font.Strike {
			font.Strike = &style.Font.Strike
		}
		if style.Font.Color != "" {
			color := "#" + strings.ToUpper(style.Font.Color)
			font.Color = &color
		}
		if style.Font.VertAlign != "" {
			va := FontVertAlign(style.Font.VertAlign)
			font.VertAlign = &va
		}
		if font.Bold != nil || font.Italic != nil || font.Underline != nil || font.Size != nil || font.Strike != nil || font.Color != nil || font.VertAlign != nil {
			result.Font = font
		}
	}
	if style.Fill.Type != "" || style.Fill.Pattern != 0 || len(style.Fill.Color) > 0 {
		fill := &FillStyle{}
		if style.Fill.Type != "" {
			fill.Type = FillType(style.Fill.Type)
		}
		if style.Fill.Pattern != 0 {
			fill.Pattern = intToFillPatternName(style.Fill.Pattern)
		}
		if len(style.Fill.Color) > 0 {
			var colors []string
			for _, color := range style.Fill.Color {
				if color != "" {
					colors = append(colors, "#"+strings.ToUpper(color))
				}
			}
			fill.Color = colors
		}
		if style.Fill.Shading != 0 {
			shading := intToFillShadingName(style.Fill.Shading)
			fill.Shading = &shading
		}
		if fill.Type != "" || fill.Pattern != FillPatternNone || len(fill.Color) > 0 || fill.Shading != nil {
			result.Fill = fill
		}
	}
	if style.CustomNumFmt != nil && *style.CustomNumFmt != "" {
		result.NumFmt = style.CustomNumFmt
	}
	if style.DecimalPlaces != nil && *style.DecimalPlaces != 0 {
		result.DecimalPlaces = style.DecimalPlaces
	}
	return result
}

func intToBorderStyleName(style int) BorderStyle {
	styles := map[int]BorderStyle{0: BorderStyleNone, 1: BorderStyleContinuous, 2: BorderStyleContinuous, 3: BorderStyleDash, 4: BorderStyleDot, 5: BorderStyleContinuous, 6: BorderStyleDouble, 7: BorderStyleContinuous, 8: BorderStyleDashDot, 9: BorderStyleDashDotDot, 10: BorderStyleSlantDashDot, 11: BorderStyleContinuous, 12: BorderStyleMediumDashDot, 13: BorderStyleMediumDashDotDot}
	if name, ok := styles[style]; ok {
		return name
	}
	return BorderStyleContinuous
}

func intToFillPatternName(pattern int) FillPattern {
	patterns := map[int]FillPattern{0: FillPatternNone, 1: FillPatternSolid, 2: FillPatternMediumGray, 3: FillPatternDarkGray, 4: FillPatternLightGray, 5: FillPatternDarkHorizontal, 6: FillPatternDarkVertical, 7: FillPatternDarkDown, 8: FillPatternDarkUp, 9: FillPatternDarkGrid, 10: FillPatternDarkTrellis, 11: FillPatternLightHorizontal, 12: FillPatternLightVertical, 13: FillPatternLightDown, 14: FillPatternLightUp, 15: FillPatternLightGrid, 16: FillPatternLightTrellis, 17: FillPatternGray125, 18: FillPatternGray0625}
	if name, ok := patterns[pattern]; ok {
		return name
	}
	return FillPatternNone
}

func intToFillShadingName(shading int) FillShading {
	shadings := map[int]FillShading{0: FillShadingHorizontal, 1: FillShadingVertical, 2: FillShadingDiagonalDown, 3: FillShadingDiagonalUp, 4: FillShadingFromCenter, 5: FillShadingFromCorner}
	if name, ok := shadings[shading]; ok {
		return name
	}
	return FillShadingHorizontal
}

func borderStyleNameToInt(style BorderStyle) int {
	styles := map[BorderStyle]int{BorderStyleNone: 0, BorderStyleContinuous: 1, BorderStyleDash: 3, BorderStyleDot: 4, BorderStyleDouble: 6, BorderStyleDashDot: 8, BorderStyleDashDotDot: 9, BorderStyleSlantDashDot: 10, BorderStyleMediumDashDot: 12, BorderStyleMediumDashDotDot: 13}
	if v, ok := styles[style]; ok {
		return v
	}
	return 1
}

func fillPatternNameToInt(pattern FillPattern) int {
	patterns := map[FillPattern]int{FillPatternNone: 0, FillPatternSolid: 1, FillPatternMediumGray: 2, FillPatternDarkGray: 3, FillPatternLightGray: 4, FillPatternDarkHorizontal: 5, FillPatternDarkVertical: 6, FillPatternDarkDown: 7, FillPatternDarkUp: 8, FillPatternDarkGrid: 9, FillPatternDarkTrellis: 10, FillPatternLightHorizontal: 11, FillPatternLightVertical: 12, FillPatternLightDown: 13, FillPatternLightUp: 14, FillPatternLightGrid: 15, FillPatternLightTrellis: 16, FillPatternGray125: 17, FillPatternGray0625: 18}
	if v, ok := patterns[pattern]; ok {
		return v
	}
	return 0
}

func fillShadingNameToInt(shading FillShading) int {
	shadings := map[FillShading]int{FillShadingHorizontal: 0, FillShadingVertical: 1, FillShadingDiagonalDown: 2, FillShadingDiagonalUp: 3, FillShadingFromCenter: 4, FillShadingFromCorner: 5}
	if v, ok := shadings[shading]; ok {
		return v
	}
	return 0
}

func (w *ExcelizeWorksheet) GetMergedCells() ([]MergedCell, error) {
	excelizeMerges, err := w.file.GetMergeCells(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get merge cells: %w", err)
	}
	result := make([]MergedCell, 0, len(excelizeMerges))
	for _, m := range excelizeMerges {
		startCol, startRow, err := excelize.CellNameToCoordinates(m.GetStartAxis())
		if err != nil {
			return nil, fmt.Errorf("failed to parse merge start axis %q: %w", m.GetStartAxis(), err)
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(m.GetEndAxis())
		if err != nil {
			return nil, fmt.Errorf("failed to parse merge end axis %q: %w", m.GetEndAxis(), err)
		}
		result = append(result, MergedCell{
			StartCol: startCol,
			StartRow: startRow,
			EndCol:   endCol,
			EndRow:   endRow,
		})
	}
	return result, nil
}

func (w *ExcelizeWorksheet) updateDimension(updatedCell string) error {
	dimension, err := w.file.GetSheetDimension(w.sheetName)
	if err != nil {
		return err
	}
	startCol, startRow, endCol, endRow, err := ParseRange(dimension)
	if err != nil {
		return err
	}
	updatedCol, updatedRow, err := excelize.CellNameToCoordinates(updatedCell)
	if err != nil {
		return err
	}
	if startCol > updatedCol {
		startCol = updatedCol
	}
	if endCol < updatedCol {
		endCol = updatedCol
	}
	if startRow > updatedRow {
		startRow = updatedRow
	}
	if endRow < updatedRow {
		endRow = updatedRow
	}
	startRange, _ := excelize.CoordinatesToCellName(startCol, startRow)
	endRange, _ := excelize.CoordinatesToCellName(endCol, endRow)
	return w.file.SetSheetDimension(w.sheetName, fmt.Sprintf("%s:%s", startRange, endRange))
}

func (w *ExcelizeWorksheet) GetValuesRange(rangeRef string) ([][]string, error) {
	values, err := w.GetValuesRangeAny(rangeRef)
	if err != nil {
		return nil, err
	}
	return anyMatrixToStringMatrix(values), nil
}

func (w *ExcelizeWorksheet) GetValuesRangeAny(rangeRef string) ([][]any, error) {
	startCol, startRow, endCol, endRow, err := ParseRange(rangeRef)
	if err != nil {
		return nil, err
	}
	numRows := endRow - startRow + 1
	numCols := endCol - startCol + 1
	result := make([][]any, numRows)
	for r := 0; r < numRows; r++ {
		result[r] = make([]any, numCols)
		for c := 0; c < numCols; c++ {
			cell, _ := excelize.CoordinatesToCellName(startCol+c, startRow+r)
			val, err := w.getValueAny(cell)
			if err != nil {
				return nil, err
			}
			result[r][c] = val
		}
	}
	return result, nil
}

func (w *ExcelizeWorksheet) GetFormulasRange(rangeRef string) ([][]string, error) {
	values, err := w.GetFormulasRangeAny(rangeRef)
	if err != nil {
		return nil, err
	}
	return anyMatrixToStringMatrix(values), nil
}

func (w *ExcelizeWorksheet) GetFormulasRangeAny(rangeRef string) ([][]any, error) {
	startCol, startRow, endCol, endRow, err := ParseRange(rangeRef)
	if err != nil {
		return nil, err
	}
	numRows := endRow - startRow + 1
	numCols := endCol - startCol + 1
	result := make([][]any, numRows)
	for r := 0; r < numRows; r++ {
		result[r] = make([]any, numCols)
		for c := 0; c < numCols; c++ {
			cell, _ := excelize.CoordinatesToCellName(startCol+c, startRow+r)
			val, err := w.getFormulaAny(cell)
			if err != nil {
				return nil, err
			}
			result[r][c] = val
		}
	}
	return result, nil
}

func (w *ExcelizeWorksheet) SetValuesRange(rangeRef string, values [][]any) error {
	startCol, startRow, _, _, err := ParseRange(rangeRef)
	if err != nil {
		return err
	}
	for rowIdx, row := range values {
		for colIdx, val := range row {
			cell, err := excelize.CoordinatesToCellName(startCol+colIdx, startRow+rowIdx)
			if err != nil {
				return err
			}
			if err := w.SetValue(cell, val); err != nil {
				return err
			}
		}
	}
	return nil
}

func (w *ExcelizeWorksheet) SetFormulasRange(rangeRef string, values [][]any) error {
	startCol, startRow, _, _, err := ParseRange(rangeRef)
	if err != nil {
		return err
	}
	for rowIdx, row := range values {
		for colIdx, val := range row {
			cell, err := excelize.CoordinatesToCellName(startCol+colIdx, startRow+rowIdx)
			if err != nil {
				return err
			}
			if strVal, ok := val.(string); ok && strings.HasPrefix(strVal, "=") {
				if err := w.SetFormula(cell, strVal); err != nil {
					return err
				}
				continue
			}
			if err := w.SetValue(cell, val); err != nil {
				return err
			}
		}
	}
	return nil
}

func (w *ExcelizeWorksheet) getValueAny(cell string) (any, error) {
	rawValue, err := w.file.GetCellValue(w.sheetName, cell, excelize.Options{RawCellValue: true})
	if err != nil {
		return nil, err
	}
	cellType, err := w.file.GetCellType(w.sheetName, cell)
	if err != nil {
		return nil, err
	}
	return coerceExcelizeCellValue(rawValue, cellType), nil
}

func (w *ExcelizeWorksheet) getFormulaAny(cell string) (any, error) {
	formula, err := w.file.GetCellFormula(w.sheetName, cell)
	if err != nil {
		return nil, fmt.Errorf("failed to get formula: %w", err)
	}
	if formula != "" {
		if !strings.HasPrefix(formula, "=") {
			formula = "=" + formula
		}
		return formula, nil
	}
	return w.getValueAny(cell)
}

func coerceExcelizeCellValue(rawValue string, cellType excelize.CellType) any {
	if rawValue == "" {
		return ""
	}
	switch cellType {
	case excelize.CellTypeBool:
		return rawValue == "1" || strings.EqualFold(rawValue, "true")
	case excelize.CellTypeNumber, excelize.CellTypeDate, excelize.CellTypeUnset:
		if number, ok := parseNumericString(rawValue); ok {
			return number
		}
		return rawValue
	default:
		return rawValue
	}
}
