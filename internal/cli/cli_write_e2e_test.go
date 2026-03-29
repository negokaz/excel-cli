package cli

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestCLIWriteEndToEnd(t *testing.T) {
	t.Parallel()

	binaryPath := buildCLIBinary(t)

	t.Run("write overwrites values in existing sheet", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createWriteCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data", "A1:C1", `[["X","Y","Z"]]`)

		if result.exitCode != 0 {
			t.Fatalf("expected exit code 0, got %d, stderr=%s", result.exitCode, result.stderr)
		}

		file, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to reopen workbook: %v", err)
		}
		defer file.Close()

		valA1, err := file.GetCellValue("Data", "A1")
		if err != nil {
			t.Fatalf("failed to get A1: %v", err)
		}
		if valA1 != "X" {
			t.Fatalf("expected X at A1, got %s", valA1)
		}

		valC1, err := file.GetCellValue("Data", "C1")
		if err != nil {
			t.Fatalf("failed to get C1: %v", err)
		}
		if valC1 != "Z" {
			t.Fatalf("expected Z at C1, got %s", valC1)
		}
	})

	t.Run("write with --newsheet creates sheet and writes values", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createWriteCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "NewSheet", "A1:B1", `[["Hello","World"]]`, "--newsheet")

		if result.exitCode != 0 {
			t.Fatalf("expected exit code 0, got %d, stderr=%s", result.exitCode, result.stderr)
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
			t.Fatalf("expected NewSheet to be created, sheets=%v", sheetList)
		}

		val, err := file.GetCellValue("NewSheet", "A1")
		if err != nil {
			t.Fatalf("failed to get A1 from NewSheet: %v", err)
		}
		if val != "Hello" {
			t.Fatalf("expected Hello at A1 in NewSheet, got %s", val)
		}
	})

	t.Run("write with --newsheet fails for existing sheet", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createWriteCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data", "A1", `[["X"]]`, "--newsheet")

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for --newsheet on existing sheet, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "already exists") {
			t.Fatalf("expected 'already exists' in error, stderr=%s", result.stderr)
		}
	})

	t.Run("write without --newsheet fails for non-existent sheet", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createWriteCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Missing", "A1", `[["X"]]`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for missing sheet without --newsheet, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "Missing") {
			t.Fatalf("expected sheet name 'Missing' in error, stderr=%s", result.stderr)
		}
	})

	t.Run("write fails when range contains sheet name", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createWriteCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data", "Data!A1:C1", `[["X","Y","Z"]]`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for range with sheet name, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "sheet name") && !strings.Contains(result.stderr, "invalid range") && !strings.Contains(result.stderr, "!") {
			t.Fatalf("expected range format error in stderr, stderr=%s", result.stderr)
		}
	})

	t.Run("write fails when values is not a JSON 2-dimensional array", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createWriteCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data", "A1:C1", `["X","Y","Z"]`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for 1D array values, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "2-dimensional") && !strings.Contains(result.stderr, "2D") && !strings.Contains(result.stderr, "2 dimensional") {
			t.Fatalf("expected 2D array error in stderr, stderr=%s", result.stderr)
		}
	})

	t.Run("write fails when values is not valid JSON", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createWriteCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data", "A1", `not-json`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for invalid JSON values, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "JSON") && !strings.Contains(result.stderr, "json") && !strings.Contains(result.stderr, "parse") {
			t.Fatalf("expected JSON parse error in stderr, stderr=%s", result.stderr)
		}
	})

	t.Run("write fails when fewer than 4 positional arguments are provided", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createWriteCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data")

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for missing arguments, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "usage") && !strings.Contains(result.stderr, "Usage") {
			t.Fatalf("expected usage hint in stderr, stderr=%s", result.stderr)
		}
	})

	t.Run("write fails when values row has more columns than range", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createWriteCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data", "A1:B1", `[["X","Y","Z"]]`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for column overflow, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "column count mismatch") {
			t.Fatalf("expected column count mismatch in stderr, stderr=%s", result.stderr)
		}
	})

	t.Run("write fails when values row has fewer columns than range", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createWriteCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data", "A1:B1", `[["X"]]`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for short column count, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "column count mismatch") {
			t.Fatalf("expected column count mismatch in stderr, stderr=%s", result.stderr)
		}
	})
}

func createWriteCLIWorkbookFixture(t *testing.T, dir string) string {
	t.Helper()

	workbook := excelize.NewFile()
	workbook.SetSheetName("Sheet1", "Data")

	mustSetCellValue(t, workbook, "Data", "A1", "OriginalA1")
	mustSetCellValue(t, workbook, "Data", "B1", "OriginalB1")
	mustSetCellValue(t, workbook, "Data", "C1", "OriginalC1")
	if err := workbook.SetSheetDimension("Data", "A1:C1"); err != nil {
		t.Fatalf("failed to set Data dimension: %v", err)
	}

	workbookPath := filepath.Join(dir, "write_fixture.xlsx")
	if err := workbook.SaveAs(workbookPath); err != nil {
		t.Fatalf("failed to save workbook: %v", err)
	}

	return workbookPath
}
