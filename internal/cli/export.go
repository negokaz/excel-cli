package cli

import (
	"flag"
	"fmt"
	"path/filepath"
	"time"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runExport(args []string) error {
	if len(args) < 1 {
		return fmt.Errorf("usage: excel-cli export <file> [<path>] --format <html|png> [options]")
	}

	fs := flag.NewFlagSet("export", flag.ContinueOnError)
	format := fs.String("format", "", "Export format")
	showFormula := fs.Bool("formula", false, "Export formulas instead of displayed values")
	showStyle := fs.Bool("style", false, "Include style information in HTML export")

	filePath := args[0]
	rawPath, remaining := extractPathArg(args[1:])
	if err := fs.Parse(remaining); err != nil {
		return err
	}
	if len(fs.Args()) != 0 {
		return fmt.Errorf("usage: excel-cli export <file> [<path>] --format <html|png> [options]")
	}
	if *format != "html" && *format != "png" {
		return fmt.Errorf("--format is missing or unsupported")
	}

	target, err := parseTargetPath(rawPath)
	if err != nil {
		return err
	}
	if target.Kind == pathKindWorkbook {
		return fmt.Errorf("unsupported path kind: %s", rawPath)
	}

	absPath, err := filepath.Abs(filePath)
	if err != nil {
		return fmt.Errorf("failed to resolve path: %w", err)
	}

	switch *format {
	case "html":
		return exportHTML(absPath, target, *showFormula, *showStyle)
	case "png":
		if *showFormula || *showStyle {
			return fmt.Errorf("--formula and --style are only supported for HTML export")
		}
		return exportPNG(absPath, target)
	default:
		return fmt.Errorf("--format is missing or unsupported")
	}
}

func exportHTML(absPath string, target targetPath, showFormula, showStyle bool) error {
	workbook, release, err := excel.OpenFile(absPath)
	if err != nil {
		return fmt.Errorf("failed to open file: %w", err)
	}
	defer release()

	worksheet, err := workbook.FindSheet(target.SheetName)
	if err != nil {
		return fmt.Errorf("sheet not found: %s", target.SheetName)
	}
	defer worksheet.Release()

	rangeRef := target.RangeRef
	if target.Kind == pathKindSheet {
		rangeRef, err = worksheet.GetDimention()
		if err != nil {
			return fmt.Errorf("failed to get sheet dimension: %w", err)
		}
		rangeRef = excel.NormalizeRange(rangeRef)
	}

	startCol, startRow, endCol, endRow, err := excel.ParseRange(rangeRef)
	if err != nil {
		return fmt.Errorf("failed to parse range: %w", err)
	}

	var tableHTML, css string
	switch {
	case showStyle && showFormula:
		tableHTML, css, err = createHTMLTableOfFormulaWithStyle(worksheet, startCol, startRow, endCol, endRow)
	case showStyle:
		tableHTML, css, err = createHTMLTableOfValuesWithStyle(worksheet, startCol, startRow, endCol, endRow)
	case showFormula:
		tableHTML, css, err = createHTMLTableOfFormula(worksheet, startCol, startRow, endCol, endRow)
	default:
		tableHTML, css, err = createHTMLTableOfValues(worksheet, startCol, startRow, endCol, endRow)
	}
	if err != nil {
		return fmt.Errorf("failed to export HTML: %w", err)
	}

	page := buildHTMLPage(HTMLPageParams{
		FilePath:      absPath,
		SheetName:     target.SheetName,
		UsedRange:     rangeRef,
		Backend:       workbook.GetBackendName(),
		GeneratedAt:   time.Now(),
		TableHTML:     tableHTML,
		StylesheetCSS: css,
	})
	outPath, err := writeOutput(page)
	if err != nil {
		return fmt.Errorf("failed to write output: %w", err)
	}
	fmt.Println(outPath)
	return nil
}

func exportPNG(absPath string, target targetPath) error {
	workbook, release, err := excel.OpenFileRequireOLE(absPath)
	if err != nil {
		return fmt.Errorf("PNG capture is not supported by the selected backend or runtime: %w", err)
	}
	defer release()

	worksheet, err := workbook.FindSheet(target.SheetName)
	if err != nil {
		return fmt.Errorf("sheet not found: %s", target.SheetName)
	}
	defer worksheet.Release()

	rangeRef := target.RangeRef
	if target.Kind == pathKindSheet {
		rangeRef, err = worksheet.GetDimention()
		if err != nil {
			return fmt.Errorf("failed to get sheet dimension: %w", err)
		}
		rangeRef = excel.NormalizeRange(rangeRef)
	}

	base64image, err := worksheet.CapturePicture(rangeRef)
	if err != nil {
		return fmt.Errorf("failed to capture picture: %w", err)
	}

	data, err := decodeBase64Image(base64image)
	if err != nil {
		return err
	}
	outPath, err := writeCaptureOutput(data)
	if err != nil {
		return fmt.Errorf("failed to write output: %w", err)
	}
	fmt.Println(outPath)
	return nil
}
