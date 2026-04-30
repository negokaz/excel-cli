package cli

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"time"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runRead(args []string) error {
	if len(args) < 2 {
		return fmt.Errorf("usage: excel-cli read <file> <sheet> [--formula] [--style]")
	}

	fs := flag.NewFlagSet("read", flag.ContinueOnError)
	showFormula := fs.Bool("formula", false, "Show formulas instead of values")
	showStyle := fs.Bool("style", false, "Include cell style information")

	filePath := args[0]
	sheetName := args[1]

	if err := fs.Parse(args[2:]); err != nil {
		return err
	}
	if len(fs.Args()) != 0 {
		return fmt.Errorf("usage: excel-cli read <file> <sheet> [--formula] [--style]")
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

	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to find sheet %q: %w", sheetName, err)
	}
	defer worksheet.Release()

	usedRange, err := worksheet.GetDimention()
	if err != nil {
		return fmt.Errorf("failed to get sheet dimension: %w", err)
	}
	isEmpty, err := excel.IsEmptyWorksheet(worksheet, usedRange)
	if err != nil {
		return fmt.Errorf("failed to inspect sheet contents: %w", err)
	}
	if isEmpty {
		return fmt.Errorf("sheet %q is empty", sheetName)
	}

	startCol, startRow, endCol, endRow, err := excel.ParseRange(usedRange)
	if err != nil {
		return fmt.Errorf("failed to parse range: %w", err)
	}

	var tableHTML, css string
	switch {
	case *showStyle && *showFormula:
		tableHTML, css, err = createHTMLTableOfFormulaWithStyle(worksheet, startCol, startRow, endCol, endRow)
	case *showStyle:
		tableHTML, css, err = createHTMLTableOfValuesWithStyle(worksheet, startCol, startRow, endCol, endRow)
	case *showFormula:
		tableHTML, css, err = createHTMLTableOfFormula(worksheet, startCol, startRow, endCol, endRow)
	default:
		tableHTML, css, err = createHTMLTableOfValues(worksheet, startCol, startRow, endCol, endRow)
	}
	if err != nil {
		return fmt.Errorf("failed to read sheet: %w", err)
	}

	page := buildHTMLPage(HTMLPageParams{
		FilePath:      absPath,
		SheetName:     sheetName,
		UsedRange:     usedRange,
		Backend:       workbook.GetBackendName(),
		GeneratedAt:   time.Now(),
		TableHTML:     tableHTML,
		StylesheetCSS: css,
	})

	outPath, err := writeOutput(page)
	if err != nil {
		return fmt.Errorf("failed to write output: %w", err)
	}

	fmt.Fprintln(os.Stdout, outPath)
	return nil
}

func outputFileName() string {
	t := time.Now().UTC()
	ms := t.Nanosecond() / 1_000_000
	return fmt.Sprintf("sheet-%s-%03dZ.html", t.Format("2006-01-02T15-04-05"), ms)
}

func writeOutput(content string) (string, error) {
	dir := ".excel-cli"
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return "", fmt.Errorf("failed to create output directory: %w", err)
	}
	filename := outputFileName()
	outPath := filepath.Join(dir, filename)
	if err := os.WriteFile(outPath, []byte(content), 0o644); err != nil {
		return "", fmt.Errorf("failed to write file: %w", err)
	}
	absPath, err := filepath.Abs(outPath)
	if err != nil {
		return outPath, nil
	}
	return absPath, nil
}
