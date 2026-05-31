package excel

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestWriteSheetWritesValuesToExistingSheet(t *testing.T) {
	t.Parallel()

	workbookPath := createWriteTestWorkbook(t)
	workbook, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	defer release()

	values := [][]any{
		{"New1", "New2", "New3"},
		{"New4", "New5", "New6"},
	}

	err = WriteSheet(workbook, "Data", "A1:C2", values, false)

	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}

	reopened, reopenedRelease, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to reopen workbook: %v", err)
	}
	defer reopenedRelease()

	sheet, err := reopened.FindSheet("Data")
	if err != nil {
		t.Fatalf("failed to find Data sheet after write: %v", err)
	}
	defer sheet.Release()

	topLeftVals, err := sheet.GetValuesRange("A1:A1")
	if err != nil {
		t.Fatalf("failed to get A1: %v", err)
	}
	if topLeftVals[0][0] != "New1" {
		t.Fatalf("expected New1, got %s", topLeftVals[0][0])
	}

	bottomRightVals, err := sheet.GetValuesRange("C2:C2")
	if err != nil {
		t.Fatalf("failed to get C2: %v", err)
	}
	if bottomRightVals[0][0] != "New6" {
		t.Fatalf("expected New6, got %s", bottomRightVals[0][0])
	}
}

func TestWriteSheetOverwritesExistingCellValues(t *testing.T) {
	t.Parallel()

	workbookPath := createWriteTestWorkbook(t)
	workbook, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	defer release()

	values := [][]any{{"NewValue"}}

	err = WriteSheet(workbook, "Data", "A1", values, false)

	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}

	reopened, reopenedRelease, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to reopen workbook: %v", err)
	}
	defer reopenedRelease()

	sheet, err := reopened.FindSheet("Data")
	if err != nil {
		t.Fatalf("failed to find Data sheet: %v", err)
	}
	defer sheet.Release()

	vals, err := sheet.GetValuesRange("A1:A1")
	if err != nil {
		t.Fatalf("failed to get A1: %v", err)
	}
	if vals[0][0] != "NewValue" {
		t.Fatalf("expected NewValue, got %s", vals[0][0])
	}
}

func TestWriteSheetStoresFormulaWhenValueStartsWithEquals(t *testing.T) {
	t.Parallel()

	workbookPath := createWriteTestWorkbook(t)
	workbook, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	defer release()

	values := [][]any{{"=SUM(1,2)"}}

	err = WriteSheet(workbook, "Data", "A1", values, false)

	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}

	reopened, reopenedRelease, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to reopen workbook: %v", err)
	}
	defer reopenedRelease()

	sheet, err := reopened.FindSheet("Data")
	if err != nil {
		t.Fatalf("failed to find Data sheet: %v", err)
	}
	defer sheet.Release()

	formulas, err := sheet.GetFormulasRange("A1:A1")
	if err != nil {
		t.Fatalf("failed to get formula A1: %v", err)
	}
	if formulas[0][0] != "=SUM(1,2)" {
		t.Fatalf("expected formula =SUM(1,2), got %s", formulas[0][0])
	}
}

func TestWriteSheetWithNewSheetCreatesSheetAndWritesValues(t *testing.T) {
	t.Parallel()

	workbookPath := createWriteTestWorkbook(t)
	workbook, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	defer release()

	values := [][]any{{"Hello", "World"}}

	err = WriteSheet(workbook, "NewSheet", "A1:B1", values, true)

	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}

	file, err := excelize.OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to reopen workbook: %v", err)
	}
	defer file.Close()

	sheetList := file.GetSheetList()
	found := false
	for _, name := range sheetList {
		if name == "NewSheet" {
			found = true
			break
		}
	}
	if !found {
		t.Fatalf("expected NewSheet to exist after write, sheets=%v", sheetList)
	}

	val, err := file.GetCellValue("NewSheet", "A1")
	if err != nil {
		t.Fatalf("failed to get A1 from NewSheet: %v", err)
	}
	if val != "Hello" {
		t.Fatalf("expected Hello, got %s", val)
	}
}

func TestWriteSheetWithNewSheetFailsForExistingSheet(t *testing.T) {
	t.Parallel()

	workbookPath := createWriteTestWorkbook(t)
	workbook, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	defer release()

	values := [][]any{{"X"}}

	err = WriteSheet(workbook, "Data", "A1", values, true)

	if err == nil {
		t.Fatal("expected error when --newsheet is specified for existing sheet")
	}
	if !strings.Contains(err.Error(), "already exists") {
		t.Fatalf("expected 'already exists' in error message, got %v", err)
	}
}

func TestWriteSheetWithoutNewSheetFailsForNonExistentSheet(t *testing.T) {
	t.Parallel()

	workbookPath := createWriteTestWorkbook(t)
	workbook, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	defer release()

	values := [][]any{{"X"}}

	err = WriteSheet(workbook, "Missing", "A1", values, false)

	if err == nil {
		t.Fatal("expected error for non-existent sheet without --newsheet")
	}
	if !strings.Contains(err.Error(), "Missing") {
		t.Fatalf("expected sheet name in error message, got %v", err)
	}
}

func TestWriteSheetFailsWhenRowCountDoesNotMatchRange(t *testing.T) {
	t.Parallel()

	workbookPath := createWriteTestWorkbook(t)
	workbook, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	defer release()

	values := [][]any{
		{"A", "B"},
		{"C", "D"},
	}

	err = WriteSheet(workbook, "Data", "A1:B3", values, false)

	if err == nil {
		t.Fatal("expected error for row count mismatch between range and values")
	}
}

func TestWriteSheetFailsWhenColumnCountExceedsRange(t *testing.T) {
	t.Parallel()

	workbookPath := createWriteTestWorkbook(t)
	workbook, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	defer release()

	values := [][]any{{"A", "B", "C"}}

	err = WriteSheet(workbook, "Data", "A1:B1", values, false)

	if err == nil {
		t.Fatal("expected error for column count overflow between range and values")
	}
	if !strings.Contains(err.Error(), "column count mismatch") {
		t.Fatalf("expected column count mismatch error, got %v", err)
	}
}

func TestWriteSheetFailsWhenColumnCountIsShorterThanRange(t *testing.T) {
	t.Parallel()

	workbookPath := createWriteTestWorkbook(t)
	workbook, release, err := OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	defer release()

	values := [][]any{{"A"}}

	err = WriteSheet(workbook, "Data", "A1:B1", values, false)

	if err == nil {
		t.Fatal("expected error for short column count between range and values")
	}
	if !strings.Contains(err.Error(), "column count mismatch") {
		t.Fatalf("expected column count mismatch error, got %v", err)
	}
}

func createWriteTestWorkbook(t *testing.T) string {
	t.Helper()

	tempDir := t.TempDir()
	workbookPath := filepath.Join(tempDir, "write_test.xlsx")

	file := excelize.NewFile()
	file.SetSheetName("Sheet1", "Data")
	if err := file.SetCellValue("Data", "A1", "OldValue"); err != nil {
		t.Fatalf("failed to set A1: %v", err)
	}
	if err := file.SaveAs(workbookPath); err != nil {
		t.Fatalf("failed to save workbook: %v", err)
	}
	if err := file.Close(); err != nil {
		t.Fatalf("failed to close workbook: %v", err)
	}

	return workbookPath
}
