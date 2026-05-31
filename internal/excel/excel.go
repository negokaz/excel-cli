package excel

import (
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

type Excel interface {
	GetBackendName() string
	GetSheets() ([]Worksheet, error)
	FindSheet(sheetName string) (Worksheet, error)
	CreateNewSheet(sheetName string) error
	CopySheet(srcSheetName, destSheetName string) error
	Save() error
}

type Worksheet interface {
	Release()
	Name() (string, error)
	GetTables() ([]Table, error)
	GetPivotTables() ([]PivotTable, error)
	GetDimention() (string, error)
	CapturePicture(captureRange string) (string, error)
	AddTable(tableRange, tableName string) error
	GetCellStyle(cell string) (*CellStyle, error)
	SetCellStyle(cell string, style *CellStyle) error
	GetMergedCells() ([]MergedCell, error)
	GetValuesRange(rangeRef string) ([][]string, error)
	GetFormulasRange(rangeRef string) ([][]string, error)
	SetValuesRange(rangeRef string, values [][]any) error
}

// MergedCell represents a merged cell region in a worksheet.
// All coordinates are 1-based (column 1 = A, row 1 = first row).
type MergedCell struct {
	StartCol int
	StartRow int
	EndCol   int
	EndRow   int
}

type Table struct {
	Name  string
	Range string
}

type PivotTable struct {
	Name  string
	Range string
}

type CellStyle struct {
	Border        []Border   `yaml:"border,omitempty" json:"border,omitempty"`
	Font          *FontStyle `yaml:"font,omitempty" json:"font,omitempty"`
	Fill          *FillStyle `yaml:"fill,omitempty" json:"fill,omitempty"`
	NumFmt        *string    `yaml:"numFmt,omitempty" json:"numFmt,omitempty"`
	DecimalPlaces *int       `yaml:"decimalPlaces,omitempty" json:"decimalPlaces,omitempty"`
}

type Border struct {
	Type  BorderType  `yaml:"type" json:"type"`
	Style BorderStyle `yaml:"style,omitempty" json:"style,omitempty"`
	Color string      `yaml:"color,omitempty" json:"color,omitempty"`
}

type FontStyle struct {
	Bold      *bool          `yaml:"bold,omitempty" json:"bold,omitempty"`
	Italic    *bool          `yaml:"italic,omitempty" json:"italic,omitempty"`
	Underline *FontUnderline `yaml:"underline,omitempty" json:"underline,omitempty"`
	Size      *int           `yaml:"size,omitempty" json:"size,omitempty"`
	Strike    *bool          `yaml:"strike,omitempty" json:"strike,omitempty"`
	Color     *string        `yaml:"color,omitempty" json:"color,omitempty"`
	VertAlign *FontVertAlign `yaml:"vertAlign,omitempty" json:"vertAlign,omitempty"`
}

type FillStyle struct {
	Type    FillType     `yaml:"type,omitempty" json:"type,omitempty"`
	Pattern FillPattern  `yaml:"pattern,omitempty" json:"pattern,omitempty"`
	Color   []string     `yaml:"color,omitempty" json:"color,omitempty"`
	Shading *FillShading `yaml:"shading,omitempty" json:"shading,omitempty"`
}

// NewFile creates a new blank Excel workbook at the specified absolute path.
// Returns an error if the file already exists.
func NewFile(absoluteFilePath string) error {
	if _, err := os.Stat(absoluteFilePath); err == nil {
		return fmt.Errorf("file already exists: %s", absoluteFilePath)
	}
	file := excelize.NewFile()
	defer file.Close()
	return file.SaveAs(absoluteFilePath)
}

// OpenFile opens an Excel file and returns an Excel interface.
// It first tries to open the file using OLE automation (Windows only),
// and if that fails, it falls back to the excelize library.
func OpenFile(absoluteFilePath string) (Excel, func(), error) {
	ole, releaseFn, err := NewExcelOle(absoluteFilePath)
	if err == nil {
		return ole, releaseFn, nil
	}
	workbook, err := excelize.OpenFile(absoluteFilePath)
	if err != nil {
		return nil, func() {}, err
	}
	ex := NewExcelizeExcel(workbook)
	return ex, func() {
		workbook.Close()
	}, nil
}

type BorderType string

const (
	BorderTypeLeft         BorderType = "left"
	BorderTypeRight        BorderType = "right"
	BorderTypeTop          BorderType = "top"
	BorderTypeBottom       BorderType = "bottom"
	BorderTypeDiagonalDown BorderType = "diagonalDown"
	BorderTypeDiagonalUp   BorderType = "diagonalUp"
)

func (b BorderType) String() string               { return string(b) }
func (b BorderType) MarshalText() ([]byte, error) { return []byte(b.String()), nil }
func BorderTypeValues() []BorderType {
	return []BorderType{BorderTypeLeft, BorderTypeRight, BorderTypeTop, BorderTypeBottom, BorderTypeDiagonalDown, BorderTypeDiagonalUp}
}

type BorderStyle string

const (
	BorderStyleNone             BorderStyle = "none"
	BorderStyleContinuous       BorderStyle = "continuous"
	BorderStyleDash             BorderStyle = "dash"
	BorderStyleDot              BorderStyle = "dot"
	BorderStyleDouble           BorderStyle = "double"
	BorderStyleDashDot          BorderStyle = "dashDot"
	BorderStyleDashDotDot       BorderStyle = "dashDotDot"
	BorderStyleSlantDashDot     BorderStyle = "slantDashDot"
	BorderStyleMediumDashDot    BorderStyle = "mediumDashDot"
	BorderStyleMediumDashDotDot BorderStyle = "mediumDashDotDot"
)

func (b BorderStyle) String() string               { return string(b) }
func (b BorderStyle) MarshalText() ([]byte, error) { return []byte(b.String()), nil }
func BorderStyleValues() []BorderStyle {
	return []BorderStyle{BorderStyleNone, BorderStyleContinuous, BorderStyleDash, BorderStyleDot, BorderStyleDouble, BorderStyleDashDot, BorderStyleDashDotDot, BorderStyleSlantDashDot, BorderStyleMediumDashDot, BorderStyleMediumDashDotDot}
}

type FontUnderline string

const (
	FontUnderlineNone             FontUnderline = "none"
	FontUnderlineSingle           FontUnderline = "single"
	FontUnderlineDouble           FontUnderline = "double"
	FontUnderlineSingleAccounting FontUnderline = "singleAccounting"
	FontUnderlineDoubleAccounting FontUnderline = "doubleAccounting"
)

func (f FontUnderline) String() string               { return string(f) }
func (f FontUnderline) MarshalText() ([]byte, error) { return []byte(f.String()), nil }
func FontUnderlineValues() []FontUnderline {
	return []FontUnderline{FontUnderlineNone, FontUnderlineSingle, FontUnderlineDouble, FontUnderlineSingleAccounting, FontUnderlineDoubleAccounting}
}

type FontVertAlign string

const (
	FontVertAlignBaseline    FontVertAlign = "baseline"
	FontVertAlignSuperscript FontVertAlign = "superscript"
	FontVertAlignSubscript   FontVertAlign = "subscript"
)

func (v FontVertAlign) String() string               { return string(v) }
func (v FontVertAlign) MarshalText() ([]byte, error) { return []byte(v.String()), nil }
func FontVertAlignValues() []FontVertAlign {
	return []FontVertAlign{FontVertAlignBaseline, FontVertAlignSuperscript, FontVertAlignSubscript}
}

type FillType string

const (
	FillTypeGradient FillType = "gradient"
	FillTypePattern  FillType = "pattern"
)

func (f FillType) String() string               { return string(f) }
func (f FillType) MarshalText() ([]byte, error) { return []byte(f.String()), nil }
func FillTypeValues() []FillType                { return []FillType{FillTypeGradient, FillTypePattern} }

type FillPattern string

const (
	FillPatternNone            FillPattern = "none"
	FillPatternSolid           FillPattern = "solid"
	FillPatternMediumGray      FillPattern = "mediumGray"
	FillPatternDarkGray        FillPattern = "darkGray"
	FillPatternLightGray       FillPattern = "lightGray"
	FillPatternDarkHorizontal  FillPattern = "darkHorizontal"
	FillPatternDarkVertical    FillPattern = "darkVertical"
	FillPatternDarkDown        FillPattern = "darkDown"
	FillPatternDarkUp          FillPattern = "darkUp"
	FillPatternDarkGrid        FillPattern = "darkGrid"
	FillPatternDarkTrellis     FillPattern = "darkTrellis"
	FillPatternLightHorizontal FillPattern = "lightHorizontal"
	FillPatternLightVertical   FillPattern = "lightVertical"
	FillPatternLightDown       FillPattern = "lightDown"
	FillPatternLightUp         FillPattern = "lightUp"
	FillPatternLightGrid       FillPattern = "lightGrid"
	FillPatternLightTrellis    FillPattern = "lightTrellis"
	FillPatternGray125         FillPattern = "gray125"
	FillPatternGray0625        FillPattern = "gray0625"
)

func (f FillPattern) String() string               { return string(f) }
func (f FillPattern) MarshalText() ([]byte, error) { return []byte(f.String()), nil }
func FillPatternValues() []FillPattern {
	return []FillPattern{FillPatternNone, FillPatternSolid, FillPatternMediumGray, FillPatternDarkGray, FillPatternLightGray, FillPatternDarkHorizontal, FillPatternDarkVertical, FillPatternDarkDown, FillPatternDarkUp, FillPatternDarkGrid, FillPatternDarkTrellis, FillPatternLightHorizontal, FillPatternLightVertical, FillPatternLightDown, FillPatternLightUp, FillPatternLightGrid, FillPatternLightTrellis, FillPatternGray125, FillPatternGray0625}
}

type FillShading string

const (
	FillShadingHorizontal   FillShading = "horizontal"
	FillShadingVertical     FillShading = "vertical"
	FillShadingDiagonalDown FillShading = "diagonalDown"
	FillShadingDiagonalUp   FillShading = "diagonalUp"
	FillShadingFromCenter   FillShading = "fromCenter"
	FillShadingFromCorner   FillShading = "fromCorner"
)

func (f FillShading) String() string               { return string(f) }
func (f FillShading) MarshalText() ([]byte, error) { return []byte(f.String()), nil }
func FillShadingValues() []FillShading {
	return []FillShading{FillShadingHorizontal, FillShadingVertical, FillShadingDiagonalDown, FillShadingDiagonalUp, FillShadingFromCenter, FillShadingFromCorner}
}
