package cli

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestCLINewEndToEnd(t *testing.T) {
	t.Parallel()

	binaryPath := buildCLIBinary(t)

	t.Run("new creates a valid xlsx workbook", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := filepath.Join(workDir, "new_workbook.xlsx")

		result := runCLICommand(t, binaryPath, workDir, "new", workbookPath)

		if result.exitCode != 0 {
			t.Fatalf("expected exit code 0, got %d, stderr=%s", result.exitCode, result.stderr)
		}

		outputPath := strings.TrimSpace(result.stdout)
		if !filepath.IsAbs(outputPath) {
			t.Fatalf("expected absolute path in stdout, got %s", outputPath)
		}

		file, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to open created workbook: %v", err)
		}
		defer file.Close()

		sheets := file.GetSheetList()
		if len(sheets) == 0 {
			t.Fatal("expected at least one sheet in the new workbook")
		}
	})

	t.Run("new fails when file already exists", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := filepath.Join(workDir, "existing.xlsx")

		// Create the file first
		first := runCLICommand(t, binaryPath, workDir, "new", workbookPath)
		if first.exitCode != 0 {
			t.Fatalf("expected first new to succeed, got %d, stderr=%s", first.exitCode, first.stderr)
		}

		result := runCLICommand(t, binaryPath, workDir, "new", workbookPath)

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code when file already exists, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "already exists") {
			t.Fatalf("expected 'already exists' in error, stderr=%s", result.stderr)
		}
	})

	t.Run("new fails when no file argument is provided", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()

		result := runCLICommand(t, binaryPath, workDir, "new")

		if result.exitCode == 0 {
			t.Fatalf("expected non-zero exit code for missing argument, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "usage") && !strings.Contains(result.stderr, "Usage") {
			t.Fatalf("expected usage hint in stderr, stderr=%s", result.stderr)
		}
	})
}
