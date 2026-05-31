//go:build windows

package excel

import (
	"bufio"
	"bytes"
	"encoding/base64"
	"fmt"
	"io"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/skanehira/clipboard-image"
	"github.com/xuri/excelize/v2"
)

type OleExcel struct {
	application         *ole.IDispatch
	workbook            *ole.IDispatch
	normalFontSize      int
	normalFontBold      bool
	normalFontItalic    bool
	normalFontColor     float64
	generalNumberFormat string
}

type OleWorksheet struct {
	excel     *OleExcel
	worksheet *ole.IDispatch
}

func NewExcelOle(absolutePath string) (Excel, func(), error) {
	runtime.LockOSThread()
	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)

	unknown, err := oleutil.GetActiveObject("Excel.Application")
	if err != nil {
		return nil, func() {}, err
	}
	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, func() {}, err
	}
	oleutil.MustPutProperty(excel, "ScreenUpdating", false)
	oleutil.MustPutProperty(excel, "EnableEvents", false)
	workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
	c := oleutil.MustGetProperty(workbooks, "Count").Val
	for i := 1; i <= int(c); i++ {
		workbook := oleutil.MustGetProperty(workbooks, "Item", i).ToIDispatch()
		fullName := oleutil.MustGetProperty(workbook, "FullName").ToString()
		name := oleutil.MustGetProperty(workbook, "Name").ToString()
		if strings.HasPrefix(fullName, "https:") && name == filepath.Base(absolutePath) {
			if FileIsNotWritable(absolutePath) {
				oleExcel, err := newOleExcel(excel, workbook)
				if err != nil {
					workbook.Release()
					return nil, func() {}, err
				}
				return oleExcel, func() {
					oleutil.MustPutProperty(excel, "EnableEvents", true)
					oleutil.MustPutProperty(excel, "ScreenUpdating", true)
					workbook.Release()
					workbooks.Release()
					excel.Release()
					ole.CoUninitialize()
					runtime.UnlockOSThread()
				}, nil
			}
		} else if normalizePath(fullName) == normalizePath(absolutePath) {
			oleExcel, err := newOleExcel(excel, workbook)
			if err != nil {
				workbook.Release()
				return nil, func() {}, err
			}
			return oleExcel, func() {
				oleutil.MustPutProperty(excel, "EnableEvents", true)
				oleutil.MustPutProperty(excel, "ScreenUpdating", true)
				workbook.Release()
				workbooks.Release()
				excel.Release()
				ole.CoUninitialize()
				runtime.UnlockOSThread()
			}, nil
		}
		workbook.Release()
	}
	return nil, func() {}, fmt.Errorf("workbook not found: %s", absolutePath)
}

func newOleExcel(application, workbook *ole.IDispatch) (*OleExcel, error) {
	normalStyleVar, err := oleutil.GetProperty(workbook, "Styles", "Normal")
	if err != nil {
		return nil, err
	}
	normalStyle := normalStyleVar.ToIDispatch()
	defer normalStyle.Release()

	normalFont := oleutil.MustGetProperty(normalStyle, "Font").ToIDispatch()
	defer normalFont.Release()

	generalNumberFormat, err := oleutil.GetProperty(application, "International", 26)
	if err != nil {
		return nil, err
	}

	return &OleExcel{
		application:         application,
		workbook:            workbook,
		normalFontSize:      int(oleutil.MustGetProperty(normalFont, "Size").Value().(float64)),
		normalFontBold:      oleutil.MustGetProperty(normalFont, "Bold").Value().(bool),
		normalFontItalic:    oleutil.MustGetProperty(normalFont, "Italic").Value().(bool),
		normalFontColor:     oleutil.MustGetProperty(normalFont, "Color").Value().(float64),
		generalNumberFormat: generalNumberFormat.Value().(string),
	}, nil
}

func (o *OleExcel) GetBackendName() string { return "ole" }

func (o *OleExcel) GetSheets() ([]Worksheet, error) {
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()
	count := int(oleutil.MustGetProperty(worksheets, "Count").Val)
	list := make([]Worksheet, count)
	for i := 1; i <= count; i++ {
		ws := oleutil.MustGetProperty(worksheets, "Item", i).ToIDispatch()
		list[i-1] = &OleWorksheet{excel: o, worksheet: ws}
	}
	return list, nil
}

func (o *OleExcel) FindSheet(sheetName string) (Worksheet, error) {
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()
	count := int(oleutil.MustGetProperty(worksheets, "Count").Val)
	for i := 1; i <= count; i++ {
		ws := oleutil.MustGetProperty(worksheets, "Item", i).ToIDispatch()
		name := oleutil.MustGetProperty(ws, "Name").ToString()
		if name == sheetName {
			return &OleWorksheet{excel: o, worksheet: ws}, nil
		}
		ws.Release()
	}
	return nil, fmt.Errorf("sheet not found: %s", sheetName)
}

func (o *OleExcel) CreateNewSheet(sheetName string) error {
	activeWorksheet := oleutil.MustGetProperty(o.workbook, "ActiveSheet").ToIDispatch()
	defer activeWorksheet.Release()
	activeWorksheetIndex := oleutil.MustGetProperty(activeWorksheet, "Index").Val
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()
	_, err := oleutil.CallMethod(worksheets, "Add", nil, activeWorksheet)
	if err != nil {
		return err
	}
	ws := oleutil.MustGetProperty(worksheets, "Item", activeWorksheetIndex+1).ToIDispatch()
	defer ws.Release()
	_, err = oleutil.PutProperty(ws, "Name", sheetName)
	return err
}

func (o *OleExcel) CopySheet(srcSheetName string, dstSheetName string) error {
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()
	srcSheetVariant, err := oleutil.GetProperty(worksheets, "Item", srcSheetName)
	if err != nil {
		return fmt.Errorf("failed to get sheet: %w", err)
	}
	srcSheet := srcSheetVariant.ToIDispatch()
	defer srcSheet.Release()
	srcSheetIndex := oleutil.MustGetProperty(srcSheet, "Index").Val
	_, err = oleutil.CallMethod(srcSheet, "Copy", nil, srcSheet)
	if err != nil {
		return err
	}
	dstSheetVariant, err := oleutil.GetProperty(worksheets, "Item", srcSheetIndex+1)
	if err != nil {
		return fmt.Errorf("failed to get copied sheet: %w", err)
	}
	dstSheet := dstSheetVariant.ToIDispatch()
	defer dstSheet.Release()
	_, err = oleutil.PutProperty(dstSheet, "Name", dstSheetName)
	return err
}

func (o *OleExcel) Save() error {
	_, err := oleutil.CallMethod(o.workbook, "Save")
	return err
}

func (o *OleWorksheet) Release() { o.worksheet.Release() }
func (o *OleWorksheet) Name() (string, error) {
	return oleutil.MustGetProperty(o.worksheet, "Name").ToString(), nil
}

func (o *OleWorksheet) GetTables() ([]Table, error) {
	tables := oleutil.MustGetProperty(o.worksheet, "ListObjects").ToIDispatch()
	defer tables.Release()
	count := int(oleutil.MustGetProperty(tables, "Count").Val)
	list := make([]Table, count)
	for i := 1; i <= count; i++ {
		table := oleutil.MustGetProperty(tables, "Item", i).ToIDispatch()
		defer table.Release()
		tableRange := oleutil.MustGetProperty(table, "Range").ToIDispatch()
		defer tableRange.Release()
		list[i-1] = Table{
			Name:  oleutil.MustGetProperty(table, "Name").ToString(),
			Range: NormalizeRange(oleutil.MustGetProperty(tableRange, "Address").ToString()),
		}
	}
	return list, nil
}

func (o *OleWorksheet) GetPivotTables() ([]PivotTable, error) {
	pivotTables := oleutil.MustGetProperty(o.worksheet, "PivotTables").ToIDispatch()
	defer pivotTables.Release()
	count := int(oleutil.MustGetProperty(pivotTables, "Count").Val)
	list := make([]PivotTable, count)
	for i := 1; i <= count; i++ {
		pt := oleutil.MustGetProperty(pivotTables, "Item", i).ToIDispatch()
		defer pt.Release()
		ptRange := oleutil.MustGetProperty(pt, "TableRange1").ToIDispatch()
		defer ptRange.Release()
		list[i-1] = PivotTable{
			Name:  oleutil.MustGetProperty(pt, "Name").ToString(),
			Range: NormalizeRange(oleutil.MustGetProperty(ptRange, "Address").ToString()),
		}
	}
	return list, nil
}

func (o *OleWorksheet) GetDimention() (string, error) {
	r := oleutil.MustGetProperty(o.worksheet, "UsedRange").ToIDispatch()
	defer r.Release()
	return NormalizeRange(oleutil.MustGetProperty(r, "Address").ToString()), nil
}

func (o *OleWorksheet) CapturePicture(captureRange string) (string, error) {
	// CopyPicture(xlScreen) captures the screen appearance, so ScreenUpdating
	// must be enabled to get a non-blank image.
	oleutil.MustPutProperty(o.excel.application, "ScreenUpdating", true)
	defer oleutil.MustPutProperty(o.excel.application, "ScreenUpdating", false)

	r := oleutil.MustGetProperty(o.worksheet, "Range", captureRange).ToIDispatch()
	defer r.Release()
	_, err := oleutil.CallMethod(r, "CopyPicture", int(1), int(2))
	if err != nil {
		return "", err
	}
	buf := new(bytes.Buffer)
	w := bufio.NewWriter(buf)
	clipReader, err := clipboard.ReadFromClipboard()
	if err != nil {
		return "", fmt.Errorf("failed to read from clipboard: %w", err)
	}
	if _, err := io.Copy(w, clipReader); err != nil {
		return "", fmt.Errorf("failed to copy clipboard data: %w", err)
	}
	if err := w.Flush(); err != nil {
		return "", fmt.Errorf("failed to flush buffer: %w", err)
	}
	return base64.StdEncoding.EncodeToString(buf.Bytes()), nil
}

func (o *OleWorksheet) AddTable(tableRange string, tableName string) error {
	tables := oleutil.MustGetProperty(o.worksheet, "ListObjects").ToIDispatch()
	defer tables.Release()
	tableVar, err := oleutil.CallMethod(tables, "Add", int(1), tableRange, nil, int(0))
	if err != nil {
		return err
	}
	table := tableVar.ToIDispatch()
	defer table.Release()
	_, err = oleutil.PutProperty(table, "Name", tableName)
	return err
}

func (o *OleWorksheet) GetCellStyle(cell string) (*CellStyle, error) {
	rng := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer rng.Release()

	style := &CellStyle{}

	font := oleutil.MustGetProperty(rng, "Font").ToIDispatch()
	defer font.Release()

	fontSize := int(oleutil.MustGetProperty(font, "Size").Value().(float64))
	fontBold := oleutil.MustGetProperty(font, "Bold").Value().(bool)
	fontItalic := oleutil.MustGetProperty(font, "Italic").Value().(bool)
	fontColor := oleutil.MustGetProperty(font, "Color").Value().(float64)

	if fontSize != o.excel.normalFontSize || fontBold != o.excel.normalFontBold || fontItalic != o.excel.normalFontItalic || fontColor != o.excel.normalFontColor {
		colorStr := bgrToRgb(fontColor)
		style.Font = &FontStyle{Bold: &fontBold, Italic: &fontItalic, Size: &fontSize, Color: &colorStr}
	}

	interior := oleutil.MustGetProperty(rng, "Interior").ToIDispatch()
	defer interior.Release()
	interiorPattern := excelPatternToFillPattern(oleutil.MustGetProperty(interior, "Pattern").Value().(int32))
	if interiorPattern != FillPatternNone {
		interiorColor := oleutil.MustGetProperty(interior, "Color").Value().(float64)
		style.Fill = &FillStyle{Type: "pattern", Pattern: interiorPattern, Color: []string{bgrToRgb(interiorColor)}}
	}

	borders := oleutil.MustGetProperty(rng, "Borders").ToIDispatch()
	defer borders.Release()
	borderPositions := []struct {
		index    int
		position BorderType
	}{{7, BorderTypeLeft}, {8, BorderTypeTop}, {9, BorderTypeBottom}, {10, BorderTypeRight}}

	bordersLineStyle := oleutil.MustGetProperty(borders, "LineStyle")
	var borderStyles []Border
	if bordersLineStyle.VT == ole.VT_NULL {
		for _, pos := range borderPositions {
			border := oleutil.MustGetProperty(borders, "Item", pos.index).ToIDispatch()
			defer border.Release()
			lineStyle := excelBorderStyleToName(oleutil.MustGetProperty(border, "LineStyle").Value().(int32))
			if lineStyle != BorderStyleNone {
				borderColor := oleutil.MustGetProperty(border, "Color").Value().(float64)
				borderStyles = append(borderStyles, Border{Type: pos.position, Style: lineStyle, Color: bgrToRgb(borderColor)})
			}
		}
	} else {
		lineStyle := excelBorderStyleToName(bordersLineStyle.Value().(int32))
		if lineStyle != BorderStyleNone {
			for _, pos := range borderPositions {
				border := oleutil.MustGetProperty(borders, "Item", pos.index).ToIDispatch()
				borderColor := oleutil.MustGetProperty(border, "Color").Value().(float64)
				borderStyles = append(borderStyles, Border{Type: pos.position, Style: lineStyle, Color: bgrToRgb(borderColor)})
			}
		}
	}
	style.Border = borderStyles

	numberFormat := oleutil.MustGetProperty(rng, "NumberFormat").ToString()
	if numberFormat != o.excel.generalNumberFormat && numberFormat != "@" {
		style.NumFmt = &numberFormat
	}
	decimalPlaces := extractDecimalPlacesFromFormat(numberFormat)
	style.DecimalPlaces = &decimalPlaces

	return style, nil
}

func (o *OleWorksheet) SetCellStyle(cell string, style *CellStyle) error {
	rng := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer rng.Release()

	if style.Font != nil {
		font := oleutil.MustGetProperty(rng, "Font").ToIDispatch()
		defer font.Release()
		if style.Font.Bold != nil {
			oleutil.PutProperty(font, "Bold", *style.Font.Bold)
		}
		if style.Font.Italic != nil {
			oleutil.PutProperty(font, "Italic", *style.Font.Italic)
		}
		if style.Font.Size != nil && *style.Font.Size > 0 {
			oleutil.PutProperty(font, "Size", *style.Font.Size)
		}
		if style.Font.Color != nil && *style.Font.Color != "" {
			oleutil.PutProperty(font, "Color", rgbToBgr(*style.Font.Color))
		}
		if style.Font.Strike != nil && *style.Font.Strike {
			oleutil.PutProperty(font, "Strikethrough", true)
		}
	}

	if style.Fill != nil {
		interior := oleutil.MustGetProperty(rng, "Interior").ToIDispatch()
		defer interior.Release()
		if style.Fill.Pattern != FillPatternNone {
			oleutil.PutProperty(interior, "Pattern", fillPatternToExcelPattern(style.Fill.Pattern))
		}
		if len(style.Fill.Color) > 0 && style.Fill.Color[0] != "" {
			oleutil.PutProperty(interior, "Color", rgbToBgr(style.Fill.Color[0]))
		}
	}

	if len(style.Border) > 0 {
		borders := oleutil.MustGetProperty(rng, "Borders").ToIDispatch()
		defer borders.Release()
		for _, borderStyle := range style.Border {
			borderIndex := borderTypeToIndex(borderStyle.Type)
			if borderIndex > 0 {
				border := oleutil.MustGetProperty(borders, "Item", borderIndex).ToIDispatch()
				defer border.Release()
				oleutil.PutProperty(border, "LineStyle", borderStyleNameToExcel(borderStyle.Style))
				if borderStyle.Color != "" {
					oleutil.PutProperty(border, "Color", rgbToBgr(borderStyle.Color))
				}
			}
		}
	}

	if style.NumFmt != nil && *style.NumFmt != "" {
		oleutil.PutProperty(rng, "NumberFormat", *style.NumFmt)
	}
	return nil
}

func (o *OleWorksheet) GetMergedCells() ([]MergedCell, error) {
	usedRangeDisp := oleutil.MustGetProperty(o.worksheet, "UsedRange").ToIDispatch()
	defer usedRangeDisp.Release()

	seen := make(map[string]struct{})
	var result []MergedCell

	rows := oleutil.MustGetProperty(usedRangeDisp, "Rows").ToIDispatch()
	defer rows.Release()
	rowCount := int(oleutil.MustGetProperty(rows, "Count").Val)

	cols := oleutil.MustGetProperty(usedRangeDisp, "Columns").ToIDispatch()
	defer cols.Release()
	colCount := int(oleutil.MustGetProperty(cols, "Count").Val)

	usedRangeAddress := NormalizeRange(oleutil.MustGetProperty(usedRangeDisp, "Address").ToString())
	startCol, startRow, _, _, err := ParseRange(usedRangeAddress)
	if err != nil {
		return nil, fmt.Errorf("failed to parse used range: %w", err)
	}

	for r := 0; r < rowCount; r++ {
		for c := 0; c < colCount; c++ {
			col := startCol + c
			row := startRow + r
			cellName, _ := excelize.CoordinatesToCellName(col, row)
			cellRng := oleutil.MustGetProperty(o.worksheet, "Range", cellName).ToIDispatch()
			isMerged := oleutil.MustGetProperty(cellRng, "MergeCells").Value()
			cellRng.Release()

			merged, ok := isMerged.(bool)
			if !ok || !merged {
				continue
			}

			cellRng2 := oleutil.MustGetProperty(o.worksheet, "Range", cellName).ToIDispatch()
			mergeArea := oleutil.MustGetProperty(cellRng2, "MergeArea").ToIDispatch()
			cellRng2.Release()
			areaAddr := NormalizeRange(oleutil.MustGetProperty(mergeArea, "Address").ToString())
			mergeArea.Release()

			if _, exists := seen[areaAddr]; exists {
				continue
			}
			seen[areaAddr] = struct{}{}

			sc, sr, ec, er, err := ParseRange(areaAddr)
			if err != nil {
				return nil, fmt.Errorf("failed to parse merge area %q: %w", areaAddr, err)
			}
			result = append(result, MergedCell{StartCol: sc, StartRow: sr, EndCol: ec, EndRow: er})
		}
	}
	return result, nil
}

func bgrToRgb(bgrColor float64) string {
	c := int32(bgrColor)
	r := (c >> 0) & 0xFF
	g := (c >> 8) & 0xFF
	b := (c >> 16) & 0xFF
	return fmt.Sprintf("#%02X%02X%02X", r, g, b)
}

func rgbToBgr(rgbColor string) int32 {
	if len(rgbColor) != 7 || rgbColor[0] != '#' {
		return 0
	}
	r := hexToByte(rgbColor[1:3])
	g := hexToByte(rgbColor[3:5])
	b := hexToByte(rgbColor[5:7])
	return int32(r) | (int32(g) << 8) | (int32(b) << 16)
}

func hexToByte(h string) byte {
	var result byte
	for _, c := range h {
		result *= 16
		switch {
		case c >= '0' && c <= '9':
			result += byte(c - '0')
		case c >= 'A' && c <= 'F':
			result += byte(c-'A') + 10
		case c >= 'a' && c <= 'f':
			result += byte(c-'a') + 10
		}
	}
	return result
}

func excelBorderStyleToName(style int32) BorderStyle {
	switch style {
	case 1:
		return BorderStyleContinuous
	case -4115:
		return BorderStyleDash
	case -4118:
		return BorderStyleDot
	case -4119:
		return BorderStyleDouble
	case 4:
		return BorderStyleDashDot
	case 5:
		return BorderStyleDashDotDot
	case 13:
		return BorderStyleSlantDashDot
	default:
		return BorderStyleNone
	}
}

func excelPatternToFillPattern(p int32) FillPattern {
	switch p {
	case 1:
		return FillPatternSolid
	case -4125:
		return FillPatternDarkGray
	case -4124:
		return FillPatternMediumGray
	case -4126:
		return FillPatternLightGray
	case -4121:
		return FillPatternGray125
	case -4127:
		return FillPatternGray0625
	case 2:
		return FillPatternDarkHorizontal
	case 3:
		return FillPatternDarkVertical
	case 4:
		return FillPatternDarkDown
	case 14:
		return FillPatternDarkUp
	case -4162:
		return FillPatternDarkGrid
	case -4166:
		return FillPatternDarkTrellis
	case 5, 9:
		return FillPatternLightHorizontal
	case 6, 12:
		return FillPatternLightVertical
	case 7, 10:
		return FillPatternLightDown
	case 8, 11:
		return FillPatternLightUp
	case 15, 16:
		return FillPatternLightGrid
	case 17, 18:
		return FillPatternLightTrellis
	default:
		return FillPatternNone
	}
}

func borderTypeToIndex(t BorderType) int {
	switch t {
	case BorderTypeLeft:
		return 7
	case BorderTypeTop:
		return 8
	case BorderTypeBottom:
		return 9
	case BorderTypeRight:
		return 10
	case BorderTypeDiagonalDown:
		return 5
	case BorderTypeDiagonalUp:
		return 6
	default:
		return 0
	}
}

func borderStyleNameToExcel(style BorderStyle) int32 {
	switch style {
	case BorderStyleContinuous:
		return 1
	case BorderStyleDash:
		return -4115
	case BorderStyleDot:
		return -4118
	case BorderStyleDouble:
		return -4119
	case BorderStyleDashDot:
		return 4
	case BorderStyleDashDotDot:
		return 5
	case BorderStyleSlantDashDot:
		return 13
	default:
		return -4142
	}
}

func fillPatternToExcelPattern(pattern FillPattern) int32 {
	switch pattern {
	case FillPatternSolid:
		return 1
	case FillPatternMediumGray:
		return -4124
	case FillPatternDarkGray:
		return -4125
	case FillPatternLightGray:
		return -4126
	case FillPatternGray125:
		return -4121
	case FillPatternGray0625:
		return -4127
	case FillPatternLightHorizontal:
		return 5
	case FillPatternLightVertical:
		return 6
	case FillPatternLightDown:
		return 7
	case FillPatternLightUp:
		return 8
	case FillPatternLightGrid:
		return 15
	case FillPatternLightTrellis:
		return 18
	case FillPatternDarkHorizontal:
		return 2
	case FillPatternDarkVertical:
		return 3
	case FillPatternDarkDown:
		return 4
	case FillPatternDarkUp:
		return 14
	case FillPatternDarkGrid:
		return -4162
	case FillPatternDarkTrellis:
		return -4166
	default:
		return -4142
	}
}

var extractDecimalPlacesRegexp = regexp.MustCompile(`\.([0#]+)`)

func extractDecimalPlacesFromFormat(format string) int {
	matches := extractDecimalPlacesRegexp.FindStringSubmatch(format)
	if len(matches) > 1 {
		return len(matches[1])
	}
	return 0
}

func normalizePath(path string) string {
	vol := filepath.VolumeName(path)
	if vol == "" {
		return path
	}
	return filepath.Clean(strings.ToUpper(vol) + path[len(vol):])
}

// readRangeAs2DStrings fetches a Range property (e.g. "Formula" or "Value2") and
// returns all cell values as a 2D row-major slice of strings.
//
// When the range is a single cell, Excel returns a scalar VARIANT instead of an array.
// That case is detected by the absence of the VT_ARRAY flag and handled as a 1×1 result.
//
// For multi-cell ranges, Excel returns a VT_ARRAY|VT_VARIANT SafeArray where
// dimension 1 is rows and dimension 2 is columns (both 1-based lower bounds).
func (o *OleWorksheet) readRangeAs2DStrings(rangeRef, property string, converter func(*ole.VARIANT) string) ([][]string, error) {
	rng := oleutil.MustGetProperty(o.worksheet, "Range", rangeRef).ToIDispatch()
	defer rng.Release()

	v, err := oleutil.GetProperty(rng, property)
	if err != nil {
		return nil, fmt.Errorf("failed to get %s for range %s: %w", property, rangeRef, err)
	}
	defer v.Clear()

	// Single-cell ranges and scalar properties return a plain VARIANT (no array flag).
	if v.VT&ole.VT_ARRAY == 0 {
		s := converter(v)
		return [][]string{{s}}, nil
	}

	sa := v.ToArray()

	lb1, err := safeArrayGetLB(sa.Array, 1)
	if err != nil {
		return nil, fmt.Errorf("safeArrayGetLBound dim1: %w", err)
	}
	ub1, err := safeArrayGetUB(sa.Array, 1)
	if err != nil {
		return nil, fmt.Errorf("safeArrayGetUBound dim1: %w", err)
	}
	lb2, err := safeArrayGetLB(sa.Array, 2)
	if err != nil {
		return nil, fmt.Errorf("safeArrayGetLBound dim2: %w", err)
	}
	ub2, err := safeArrayGetUB(sa.Array, 2)
	if err != nil {
		return nil, fmt.Errorf("safeArrayGetUBound dim2: %w", err)
	}

	nRows := int(ub1 - lb1 + 1)
	nCols := int(ub2 - lb2 + 1)

	result := make([][]string, nRows)
	for i := range result {
		result[i] = make([]string, nCols)
		for j := range result[i] {
			elem, err := safeArray2DGetVariant(sa.Array, lb1+int32(i), lb2+int32(j))
			if err != nil {
				result[i][j] = ""
				continue
			}
			result[i][j] = converter(&elem)
			elem.Clear()
		}
	}
	return result, nil
}

// GetValuesRange reads all cell values in rangeRef using Range.Value2 (bulk COM call).
// String and boolean values are returned exactly. Numeric values are rendered as their
// shortest exact decimal representation (no locale-specific number formatting is applied).
func (o *OleWorksheet) GetValuesRange(rangeRef string) ([][]string, error) {
	return o.readRangeAs2DStrings(rangeRef, "Value2", variantToValueString)
}

// GetFormulasRange reads all cell formulas in rangeRef using Range.Formula (bulk COM call).
// Formula cells return the formula string (e.g. "=SUM(A1:A10)").
// Non-formula cells return the cell's value as a string.
func (o *OleWorksheet) GetFormulasRange(rangeRef string) ([][]string, error) {
	return o.readRangeAs2DStrings(rangeRef, "Formula", variantToFormulaString)
}

// SetValuesRange writes all values in a single bulk COM call via a 2D VT_VARIANT SafeArray.
// Range.Formula is always used: strings beginning with "=" are evaluated as formulas,
// while all other values are treated as literals by Excel.
func (o *OleWorksheet) SetValuesRange(rangeRef string, values [][]any) error {
	if len(values) == 0 {
		return nil
	}
	numRows := len(values)
	numCols := len(values[0])
	if numCols == 0 {
		return nil
	}

	sa, err := safeArrayCreate2DVariant(numRows, numCols)
	if err != nil {
		return fmt.Errorf("failed to create SafeArray: %w", err)
	}

	for rowIdx, row := range values {
		for colIdx, val := range row {
			elem := valueToVariant(val)
			putErr := safeArray2DPutVariant(sa, int32(rowIdx+1), int32(colIdx+1), &elem)
			elem.Clear()
			if putErr != nil {
				return fmt.Errorf("failed to set element [%d][%d]: %w", rowIdx, colIdx, putErr)
			}
		}
	}

	rng := oleutil.MustGetProperty(o.worksheet, "Range", rangeRef).ToIDispatch()
	defer rng.Release()

	// Wrap SafeArray in a VARIANT. Passing *ole.VARIANT to oleutil.PutProperty
	// results in a VT_BYREF|VT_VARIANT wrapper, which Excel dereferences correctly.
	arrayVariant := ole.NewVariant(ole.VT_ARRAY|ole.VT_VARIANT, int64(uintptr(unsafe.Pointer(sa))))
	_, err = oleutil.PutProperty(rng, "Formula", &arrayVariant)
	if err != nil {
		return fmt.Errorf("failed to set range formula: %w", err)
	}
	return nil
}
