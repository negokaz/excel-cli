package excel

import (
	"path/filepath"
	"runtime"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestOpenFileOpensWorkbookAndExposesBackend(t *testing.T) {
	t.Parallel()

	workbookPath := createExcelizeTestWorkbook(t)

	workbook, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("expected workbook to open, got %v", err)
	}
	defer release()

	backend := workbook.GetBackendName()
	if backend == "" {
		t.Fatal("expected backend name")
	}
	if runtime.GOOS != "windows" && backend != "excelize" {
		t.Fatalf("expected excelize backend on non-Windows, got %s", backend)
	}
}

func TestExcelizeWorkbookSupportsReadingAndWriting(t *testing.T) {
	t.Parallel()

	workbookPath := createExcelizeTestWorkbook(t)
	file, err := excelize.OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	t.Cleanup(func() { _ = file.Close() })
	workbook := &ExcelizeExcel{file: file}

	sheets, err := workbook.GetSheets()
	if err != nil {
		t.Fatalf("expected sheets, got %v", err)
	}
	if len(sheets) != 2 {
		t.Fatalf("expected 2 sheets, got %d", len(sheets))
	}

	worksheet, err := workbook.FindSheet("Data")
	if err != nil {
		t.Fatalf("expected Data sheet, got %v", err)
	}

	values, err := worksheet.GetValuesRange("A1:A1")
	if err != nil {
		t.Fatalf("expected values, got %v", err)
	}
	formulas, err := worksheet.GetFormulasRange("B1:B1")
	if err != nil {
		t.Fatalf("expected formulas, got %v", err)
	}

	err = worksheet.SetValuesRange("C3:C3", [][]any{{"new"}})
	if err != nil {
		t.Fatalf("expected SetValuesRange to succeed, got %v", err)
	}
	err = worksheet.SetValuesRange("D4:D4", [][]any{{"=SUM(2,3)"}})
	if err != nil {
		t.Fatalf("expected SetValuesRange (formula) to succeed, got %v", err)
	}
	dimension, err := worksheet.GetDimention()
	if err != nil {
		t.Fatalf("expected dimension, got %v", err)
	}

	if values[0][0] != "plain" {
		t.Fatalf("expected plain, got %s", values[0][0])
	}
	if formulas[0][0] != "=SUM(1,2)" {
		t.Fatalf("expected =SUM(1,2), got %s", formulas[0][0])
	}
	if dimension != "A1:D4" {
		t.Fatalf("expected A1:D4, got %s", dimension)
	}
}

func TestExcelizeWorksheetRoundTripsStyleAndSave(t *testing.T) {
	t.Parallel()

	workbookPath := createExcelizeTestWorkbook(t)
	file, err := excelize.OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	t.Cleanup(func() { _ = file.Close() })
	workbook := &ExcelizeExcel{file: file}
	worksheet, err := workbook.FindSheet("Data")
	if err != nil {
		t.Fatalf("expected Data sheet, got %v", err)
	}

	bold := true
	color := "#00AA00"
	numFmt := "0.00"
	decimalPlaces := 2
	style := &CellStyle{
		Border:        []Border{{Type: BorderTypeLeft, Style: BorderStyleContinuous, Color: "#FF0000"}},
		Font:          &FontStyle{Bold: &bold, Color: &color},
		Fill:          &FillStyle{Type: FillTypePattern, Pattern: FillPatternSolid, Color: []string{"#FFF2CC"}},
		NumFmt:        &numFmt,
		DecimalPlaces: &decimalPlaces,
	}

	if err := worksheet.SetCellStyle("C2", style); err != nil {
		t.Fatalf("expected SetCellStyle to succeed, got %v", err)
	}
	gotStyle, err := worksheet.GetCellStyle("C2")
	if err != nil {
		t.Fatalf("expected GetCellStyle to succeed, got %v", err)
	}
	if err := workbook.Save(); err != nil {
		t.Fatalf("expected Save to succeed, got %v", err)
	}

	reopened, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("expected reopened workbook, got %v", err)
	}
	defer release()
	reopenedSheet, err := reopened.FindSheet("Data")
	if err != nil {
		t.Fatalf("expected reopened Data sheet, got %v", err)
	}
	reopenedStyle, err := reopenedSheet.GetCellStyle("C2")
	if err != nil {
		t.Fatalf("expected reopened style, got %v", err)
	}

	if gotStyle == nil || gotStyle.Font == nil || gotStyle.Font.Bold == nil || !*gotStyle.Font.Bold {
		t.Fatalf("expected bold style after set, got %+v", gotStyle)
	}
	if gotStyle.Fill == nil || len(gotStyle.Fill.Color) == 0 || gotStyle.Fill.Color[0] != "#FFF2CC" {
		t.Fatalf("expected fill color after set, got %+v", gotStyle)
	}
	if gotStyle.NumFmt == nil || *gotStyle.NumFmt != "0.00" {
		t.Fatalf("expected number format after set, got %+v", gotStyle)
	}
	if reopenedStyle == nil || reopenedStyle.DecimalPlaces == nil || *reopenedStyle.DecimalPlaces != 2 {
		t.Fatalf("expected decimal places after reopen, got %+v", reopenedStyle)
	}
}

func TestExcelizeWorksheetCapturePictureReturnsUnsupportedError(t *testing.T) {
	t.Parallel()

	workbookPath := createExcelizeTestWorkbook(t)
	file, err := excelize.OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	t.Cleanup(func() { _ = file.Close() })
	worksheet := &ExcelizeWorksheet{file: file, sheetName: "Data"}

	_, err = worksheet.CapturePicture("A1:B2")

	if err == nil {
		t.Fatal("expected unsupported error")
	}
	if !strings.Contains(err.Error(), "not supported") {
		t.Fatalf("expected unsupported error, got %v", err)
	}
}

func TestExcelizeWorksheetGetMergedCellsReturnsMergedAreas(t *testing.T) {
	t.Parallel()

	tempDir := t.TempDir()
	workbookPath := filepath.Join(tempDir, "merge-test.xlsx")
	file := excelize.NewFile()
	if err := file.MergeCell("Sheet1", "A1", "C2"); err != nil {
		t.Fatalf("failed to merge cells: %v", err)
	}
	if err := file.SaveAs(workbookPath); err != nil {
		t.Fatalf("failed to save: %v", err)
	}
	_ = file.Close()

	reopened, err := excelize.OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open: %v", err)
	}
	t.Cleanup(func() { _ = reopened.Close() })

	ws := &ExcelizeWorksheet{file: reopened, sheetName: "Sheet1"}
	merges, err := ws.GetMergedCells()
	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}
	if len(merges) != 1 {
		t.Fatalf("expected 1 merged area, got %d", len(merges))
	}
	m := merges[0]
	if m.StartCol != 1 || m.StartRow != 1 || m.EndCol != 3 || m.EndRow != 2 {
		t.Fatalf("unexpected merge area: %+v", m)
	}
}

func createExcelizeTestWorkbook(t *testing.T) string {
	t.Helper()

	tempDir := t.TempDir()
	workbookPath := filepath.Join(tempDir, "workbook.xlsx")

	file := excelize.NewFile()
	file.SetSheetName("Sheet1", "Data")
	if _, err := file.NewSheet("Extra"); err != nil {
		t.Fatalf("failed to create Extra sheet: %v", err)
	}
	if err := file.SetCellValue("Data", "A1", "plain"); err != nil {
		t.Fatalf("failed to set A1: %v", err)
	}
	if err := file.SetCellFormula("Data", "B1", "=SUM(1,2)"); err != nil {
		t.Fatalf("failed to set B1 formula: %v", err)
	}
	if err := file.SaveAs(workbookPath); err != nil {
		t.Fatalf("failed to save workbook: %v", err)
	}

	return workbookPath
}
