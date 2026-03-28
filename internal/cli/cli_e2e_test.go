package cli

import (
	"encoding/json"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestCLIEndToEnd(t *testing.T) {
	t.Parallel()

	binaryPath := buildCLIBinary(t)

	t.Run("list returns sheet metadata as JSON", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "list", workbookPath)

		if result.exitCode != 0 {
			t.Fatalf("expected exit code 0, got %d, stderr=%s", result.exitCode, result.stderr)
		}

		var payload struct {
			Backend string `json:"backend"`
			Sheets  []struct {
				Name      string `json:"name"`
				UsedRange string `json:"usedRange"`
			} `json:"sheets"`
		}
		if err := json.Unmarshal([]byte(result.stdout), &payload); err != nil {
			t.Fatalf("expected JSON output, got error: %v\nstdout=%s", err, result.stdout)
		}

		if payload.Backend == "" {
			t.Fatal("expected backend to be present")
		}
		if payload.Backend != "excelize" && payload.Backend != "ole" {
			t.Fatalf("unexpected backend: %s", payload.Backend)
		}
		if len(payload.Sheets) != 2 {
			t.Fatalf("expected 2 sheets, got %d", len(payload.Sheets))
		}
		if payload.Sheets[0].Name != "Data" || payload.Sheets[0].UsedRange != "A1:C2" {
			t.Fatalf("unexpected first sheet: %+v", payload.Sheets[0])
		}
		if payload.Sheets[1].Name != "Empty" || payload.Sheets[1].UsedRange != "A1" {
			t.Fatalf("unexpected second sheet: %+v", payload.Sheets[1])
		}
	})

	t.Run("read writes HTML file and prints absolute path", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "Data")

		if result.exitCode != 0 {
			t.Fatalf("expected exit code 0, got %d, stderr=%s", result.exitCode, result.stderr)
		}

		outputPath := strings.TrimSpace(result.stdout)
		if !filepath.IsAbs(outputPath) {
			t.Fatalf("expected absolute path, got %s", outputPath)
		}
		if filepath.Dir(outputPath) != filepath.Join(workDir, ".excel-cli") {
			t.Fatalf("expected output in .excel-cli, got %s", outputPath)
		}

		content, err := os.ReadFile(outputPath)
		if err != nil {
			t.Fatalf("failed to read HTML output: %v", err)
		}

		html := string(content)
		if !strings.Contains(html, "<h1>Data</h1>") {
			t.Fatalf("expected sheet header in HTML: %s", html)
		}
		if !strings.Contains(html, "<th>A</th><th>B</th><th>C</th>") {
			t.Fatalf("expected column headers in HTML: %s", html)
		}
		if !strings.Contains(html, "Line 1<br>Line 2") {
			t.Fatalf("expected newline conversion in HTML: %s", html)
		}
	})

	t.Run("read supports formula and style flags after positional arguments", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		formulaResult := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "Data", "--formula")
		if formulaResult.exitCode != 0 {
			t.Fatalf("expected formula mode to succeed, got %d, stderr=%s", formulaResult.exitCode, formulaResult.stderr)
		}
		formulaHTML := readCLIHTML(t, strings.TrimSpace(formulaResult.stdout))
		if !strings.Contains(formulaHTML, "=SUM(1,2)") {
			t.Fatalf("expected formula output in HTML: %s", formulaHTML)
		}

		styleResult := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "Data", "--style")
		if styleResult.exitCode != 0 {
			t.Fatalf("expected style mode to succeed, got %d, stderr=%s", styleResult.exitCode, styleResult.stderr)
		}
		styleHTML := readCLIHTML(t, strings.TrimSpace(styleResult.stdout))
		if !strings.Contains(styleHTML, "Style Definitions") {
			t.Fatalf("expected style definitions in HTML: %s", styleHTML)
		}
		if !strings.Contains(styleHTML, `style-ref="`) {
			t.Fatalf("expected style-ref attribute in HTML: %s", styleHTML)
		}

		bothResult := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "Data", "--formula", "--style")
		if bothResult.exitCode != 0 {
			t.Fatalf("expected formula+style mode to succeed, got %d, stderr=%s", bothResult.exitCode, bothResult.stderr)
		}
		bothHTML := readCLIHTML(t, strings.TrimSpace(bothResult.stdout))
		if !strings.Contains(bothHTML, "=SUM(1,2)") || !strings.Contains(bothHTML, `style-ref="`) {
			t.Fatalf("expected combined formula/style output: %s", bothHTML)
		}
	})

	t.Run("read treats flag-like sheet names as positional arguments", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookWithFlagNamedSheet(t, workDir)

		styleSheetResult := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "--style")
		if styleSheetResult.exitCode != 0 {
			t.Fatalf("expected --style sheet name to succeed, got %d, stderr=%s", styleSheetResult.exitCode, styleSheetResult.stderr)
		}
		styleSheetHTML := readCLIHTML(t, strings.TrimSpace(styleSheetResult.stdout))
		if !strings.Contains(styleSheetHTML, "<h1>--style</h1>") || !strings.Contains(styleSheetHTML, "style sheet value") {
			t.Fatalf("expected --style sheet contents in HTML: %s", styleSheetHTML)
		}

		formulaSheetResult := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "--formula")
		if formulaSheetResult.exitCode != 0 {
			t.Fatalf("expected --formula sheet name to succeed, got %d, stderr=%s", formulaSheetResult.exitCode, formulaSheetResult.stderr)
		}
		formulaSheetHTML := readCLIHTML(t, strings.TrimSpace(formulaSheetResult.stdout))
		if !strings.Contains(formulaSheetHTML, "<h1>--formula</h1>") || !strings.Contains(formulaSheetHTML, "formula sheet value") {
			t.Fatalf("expected --formula sheet contents in HTML: %s", formulaSheetHTML)
		}
	})

	t.Run("read fails for missing sheet and empty sheet", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		missingSheet := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "Missing")
		if missingSheet.exitCode == 0 {
			t.Fatalf("expected missing sheet to fail, stdout=%s", missingSheet.stdout)
		}
		if !strings.Contains(missingSheet.stderr, `failed to find sheet "Missing"`) {
			t.Fatalf("expected missing sheet error, stderr=%s", missingSheet.stderr)
		}

		emptySheet := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "Empty")
		if emptySheet.exitCode == 0 {
			t.Fatalf("expected empty sheet to fail, stdout=%s", emptySheet.stdout)
		}
		if !strings.Contains(emptySheet.stderr, `sheet "Empty" is empty`) {
			t.Fatalf("expected empty sheet error, stderr=%s", emptySheet.stderr)
		}
	})

	t.Run("invalid commands and missing arguments fail with usage", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		listResult := runCLICommand(t, binaryPath, workDir, "list")
		if listResult.exitCode == 0 {
			t.Fatalf("expected list without file to fail, stdout=%s", listResult.stdout)
		}
		if !strings.Contains(listResult.stderr, "usage: excel-cli list <file>") {
			t.Fatalf("expected list usage error, stderr=%s", listResult.stderr)
		}

		readResult := runCLICommand(t, binaryPath, workDir, "read", workbookPath)
		if readResult.exitCode == 0 {
			t.Fatalf("expected read without sheet to fail, stdout=%s", readResult.stdout)
		}
		if !strings.Contains(readResult.stderr, "usage: excel-cli read <file> <sheet> [--formula] [--style]") {
			t.Fatalf("expected read usage error, stderr=%s", readResult.stderr)
		}

		unknownResult := runCLICommand(t, binaryPath, workDir, "unknown")
		if unknownResult.exitCode == 0 {
			t.Fatalf("expected unknown command to fail, stdout=%s", unknownResult.stdout)
		}
		if !strings.Contains(unknownResult.stderr, "unknown command: unknown") {
			t.Fatalf("expected unknown command error, stderr=%s", unknownResult.stderr)
		}
		if !strings.Contains(unknownResult.stderr, "Usage:") {
			t.Fatalf("expected usage text for unknown command, stderr=%s", unknownResult.stderr)
		}
	})
}

type cliRunResult struct {
	stdout   string
	stderr   string
	exitCode int
}

func buildCLIBinary(t *testing.T) string {
	t.Helper()

	repoRoot := projectRoot(t)
	tempDir := t.TempDir()
	binaryName := "excel-cli-test"
	if runtime.GOOS == "windows" {
		binaryName += ".exe"
	}
	binaryPath := filepath.Join(tempDir, binaryName)
	goCache := filepath.Join(tempDir, "gocache")
	goTmp := filepath.Join(tempDir, "gotmp")
	if err := os.MkdirAll(goCache, 0o755); err != nil {
		t.Fatalf("failed to create GOCACHE: %v", err)
	}
	if err := os.MkdirAll(goTmp, 0o755); err != nil {
		t.Fatalf("failed to create GOTMPDIR: %v", err)
	}

	buildCmd := exec.Command("go", "build", "-o", binaryPath, ".")
	buildCmd.Dir = repoRoot
	buildCmd.Env = append(os.Environ(),
		"GOCACHE="+goCache,
		"GOTMPDIR="+goTmp,
	)
	output, err := buildCmd.CombinedOutput()
	if err != nil {
		t.Fatalf("failed to build CLI binary: %v\n%s", err, string(output))
	}

	return binaryPath
}

func createCLIWorkbookFixture(t *testing.T, dir string) string {
	t.Helper()

	workbook := excelize.NewFile()
	workbook.SetSheetName("Sheet1", "Data")
	if _, err := workbook.NewSheet("Empty"); err != nil {
		t.Fatalf("failed to create empty sheet: %v", err)
	}

	mustSetCellValue(t, workbook, "Data", "A1", "Name")
	mustSetCellValue(t, workbook, "Data", "B1", "Formula")
	mustSetCellValue(t, workbook, "Data", "C1", "Styled")
	mustSetCellValue(t, workbook, "Data", "A2", "Line 1\nLine 2")
	mustSetCellFormula(t, workbook, "Data", "B2", "=SUM(1,2)")
	mustSetCellValue(t, workbook, "Data", "C2", "Styled cell")
	if err := workbook.SetSheetDimension("Data", "A1:C2"); err != nil {
		t.Fatalf("failed to set Data dimension: %v", err)
	}
	if err := workbook.SetSheetDimension("Empty", "A1"); err != nil {
		t.Fatalf("failed to set Empty dimension: %v", err)
	}

	styleID, err := workbook.NewStyle(&excelize.Style{
		Border: []excelize.Border{{Type: "left", Style: 1, Color: "FF0000"}},
		Font:   &excelize.Font{Bold: true, Color: "00AA00"},
		Fill:   excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"FFF2CC"}},
	})
	if err != nil {
		t.Fatalf("failed to create style: %v", err)
	}
	if err := workbook.SetCellStyle("Data", "C2", "C2", styleID); err != nil {
		t.Fatalf("failed to apply style: %v", err)
	}

	workbookPath := filepath.Join(dir, "fixture.xlsx")
	if err := workbook.SaveAs(workbookPath); err != nil {
		t.Fatalf("failed to save workbook: %v", err)
	}

	return workbookPath
}

func createCLIWorkbookWithFlagNamedSheet(t *testing.T, dir string) string {
	t.Helper()

	workbook := excelize.NewFile()
	workbook.SetSheetName("Sheet1", "--style")
	formulaSheetIndex, err := workbook.NewSheet("--formula")
	if err != nil {
		t.Fatalf("failed to create --formula sheet: %v", err)
	}
	workbook.SetActiveSheet(formulaSheetIndex)

	mustSetCellValue(t, workbook, "--style", "A1", "style sheet value")
	mustSetCellValue(t, workbook, "--formula", "A1", "formula sheet value")
	if err := workbook.SetSheetDimension("--style", "A1"); err != nil {
		t.Fatalf("failed to set --style dimension: %v", err)
	}
	if err := workbook.SetSheetDimension("--formula", "A1"); err != nil {
		t.Fatalf("failed to set --formula dimension: %v", err)
	}

	workbookPath := filepath.Join(dir, "flag-named-sheet.xlsx")
	if err := workbook.SaveAs(workbookPath); err != nil {
		t.Fatalf("failed to save workbook: %v", err)
	}

	return workbookPath
}

func mustSetCellValue(t *testing.T, workbook *excelize.File, sheet, cell, value string) {
	t.Helper()
	if err := workbook.SetCellValue(sheet, cell, value); err != nil {
		t.Fatalf("failed to set %s: %v", cell, err)
	}
}

func mustSetCellFormula(t *testing.T, workbook *excelize.File, sheet, cell, formula string) {
	t.Helper()
	if err := workbook.SetCellFormula(sheet, cell, formula); err != nil {
		t.Fatalf("failed to set formula %s: %v", cell, err)
	}
}

func runCLICommand(t *testing.T, binaryPath, workDir string, args ...string) cliRunResult {
	t.Helper()

	cmd := exec.Command(binaryPath, args...)
	cmd.Dir = workDir
	output, err := cmd.CombinedOutput()
	text := string(output)
	if err == nil {
		return cliRunResult{stdout: text, exitCode: 0}
	}

	exitCode := 1
	if exitErr, ok := err.(*exec.ExitError); ok {
		exitCode = exitErr.ExitCode()
	}
	return cliRunResult{stderr: text, exitCode: exitCode}
}

func readCLIHTML(t *testing.T, path string) string {
	t.Helper()

	content, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("failed to read HTML file %s: %v", path, err)
	}
	return string(content)
}

func projectRoot(t *testing.T) string {
	t.Helper()

	root, err := filepath.Abs(filepath.Join("..", ".."))
	if err != nil {
		t.Fatalf("failed to resolve project root: %v", err)
	}
	return root
}
