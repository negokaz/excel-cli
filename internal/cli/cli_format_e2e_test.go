package cli

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestCLIFormatEndToEnd(t *testing.T) {
	t.Parallel()

	binaryPath := buildCLIBinary(t)

	t.Run("format applies font bold style to a cell", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createFormatCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Data", "B2:B2", `[[{"font":{"bold":true}}]]`)

		if result.exitCode != 0 {
			t.Fatalf("expected exit code 0, got %d, stderr=%s", result.exitCode, result.stderr)
		}

		file, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to reopen workbook: %v", err)
		}
		defer file.Close()

		styleID, err := file.GetCellStyle("Data", "B2")
		if err != nil {
			t.Fatalf("failed to get style for B2: %v", err)
		}
		style, err := file.GetStyle(styleID)
		if err != nil {
			t.Fatalf("failed to get style object: %v", err)
		}
		if style.Font == nil || !style.Font.Bold {
			t.Fatalf("expected B2 to have bold font, got style=%+v", style)
		}
	})

	t.Run("format applies fill color to a range", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createFormatCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Data", "A1:B1",
			`[[{"fill":{"type":"pattern","pattern":"solid","color":["#FFFF00"]}},{"fill":{"type":"pattern","pattern":"solid","color":["#FFFF00"]}}]]`)

		if result.exitCode != 0 {
			t.Fatalf("expected exit code 0, got %d, stderr=%s", result.exitCode, result.stderr)
		}

		file, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to reopen workbook: %v", err)
		}
		defer file.Close()

		for _, cell := range []string{"A1", "B1"} {
			styleID, err := file.GetCellStyle("Data", cell)
			if err != nil {
				t.Fatalf("failed to get style for %s: %v", cell, err)
			}
			style, err := file.GetStyle(styleID)
			if err != nil {
				t.Fatalf("failed to get style object for %s: %v", cell, err)
			}
			if style.Fill.Type != "pattern" {
				t.Fatalf("expected pattern fill for %s, got %+v", cell, style.Fill)
			}
		}
	})

	t.Run("format skips null cells", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createFormatCLIWorkbookFixture(t, workDir)

		// Apply bold to A1 first
		applyResult := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Data", "A1:A1", `[[{"font":{"bold":true}}]]`)
		if applyResult.exitCode != 0 {
			t.Fatalf("initial format failed: %s", applyResult.stderr)
		}

		// Now format A1:B1 with null for A1 — A1 bold should remain untouched
		result := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Data", "A1:B1", `[[null,{"font":{"italic":true}}]]`)
		if result.exitCode != 0 {
			t.Fatalf("expected exit code 0, got %d, stderr=%s", result.exitCode, result.stderr)
		}

		file, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to reopen workbook: %v", err)
		}
		defer file.Close()

		styleID, err := file.GetCellStyle("Data", "A1")
		if err != nil {
			t.Fatalf("failed to get A1 style: %v", err)
		}
		style, err := file.GetStyle(styleID)
		if err != nil {
			t.Fatalf("failed to get A1 style object: %v", err)
		}
		if style.Font == nil || !style.Font.Bold {
			t.Fatalf("expected A1 to retain bold font after null skip, got %+v", style)
		}
	})

	t.Run("format fails when row count does not match range", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createFormatCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Data", "A1:B2", `[[null,null]]`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for row count mismatch, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "row count mismatch") {
			t.Fatalf("expected 'row count mismatch' in stderr, got: %s", result.stderr)
		}
	})

	t.Run("format fails when column count does not match range", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createFormatCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Data", "A1:C1", `[[null,null]]`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for column count mismatch, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "column count mismatch") {
			t.Fatalf("expected 'column count mismatch' in stderr, got: %s", result.stderr)
		}
	})

	t.Run("format fails for missing sheet", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createFormatCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Missing", "A1:A1", `[[null]]`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for missing sheet, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "Missing") {
			t.Fatalf("expected sheet name 'Missing' in error, stderr=%s", result.stderr)
		}
	})

	t.Run("format fails for invalid JSON styles", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createFormatCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Data", "A1:A1", `not-json`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for invalid JSON, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "JSON") && !strings.Contains(result.stderr, "json") && !strings.Contains(result.stderr, "parse") {
			t.Fatalf("expected JSON parse error in stderr, got: %s", result.stderr)
		}
	})

	t.Run("format fails when range contains sheet name", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createFormatCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Data", "Data!A1:A1", `[[null]]`)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for range with sheet name, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "sheet name") && !strings.Contains(result.stderr, "!") {
			t.Fatalf("expected range format error in stderr, got: %s", result.stderr)
		}
	})

	t.Run("format fails when fewer than 4 arguments are provided", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createFormatCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Data")

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for missing arguments, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "usage") && !strings.Contains(result.stderr, "Usage") {
			t.Fatalf("expected usage hint in stderr, got: %s", result.stderr)
		}
	})

	t.Run("format fails when more than 4 arguments are provided", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createFormatCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "format", workbookPath, "Data", "A1:A1", `[[null]]`, "extra")

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for extra arguments, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "usage") && !strings.Contains(result.stderr, "Usage") {
			t.Fatalf("expected usage hint in stderr, got: %s", result.stderr)
		}
	})
}

func createFormatCLIWorkbookFixture(t *testing.T, dir string) string {
	t.Helper()

	workbook := excelize.NewFile()
	workbook.SetSheetName("Sheet1", "Data")

	mustSetCellValue(t, workbook, "Data", "A1", "Header1")
	mustSetCellValue(t, workbook, "Data", "B1", "Header2")
	mustSetCellValue(t, workbook, "Data", "A2", "Value1")
	mustSetCellValue(t, workbook, "Data", "B2", "Value2")
	if err := workbook.SetSheetDimension("Data", "A1:B2"); err != nil {
		t.Fatalf("failed to set Data dimension: %v", err)
	}

	workbookPath := filepath.Join(dir, "format_fixture.xlsx")
	if err := workbook.SaveAs(workbookPath); err != nil {
		t.Fatalf("failed to save workbook: %v", err)
	}

	return workbookPath
}
