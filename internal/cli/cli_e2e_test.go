package cli

import (
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

func TestCLIDesignEndToEnd(t *testing.T) {
	t.Parallel()

	binaryPath := buildCLIBinary(t)

	t.Run("query enumerates sheets at workbook root", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "query", workbookPath, "")
		if result.exitCode != 0 {
			t.Fatalf("expected success, stderr=%s", result.stderr)
		}

		var payload struct {
			Path    string `json:"path"`
			Kind    string `json:"kind"`
			Backend string `json:"backend"`
			Items   []struct {
				Path string `json:"path"`
				Kind string `json:"kind"`
				Name string `json:"name"`
			} `json:"items"`
		}
		if err := json.Unmarshal([]byte(result.stdout), &payload); err != nil {
			t.Fatalf("expected JSON output: %v\n%s", err, result.stdout)
		}
		if payload.Path != "" || payload.Kind != "sheetCollection" {
			t.Fatalf("unexpected payload: %+v", payload)
		}
		if len(payload.Items) != 3 {
			t.Fatalf("expected 3 sheets, got %d", len(payload.Items))
		}
		if payload.Items[1].Path != "Hidden%20Sheet" {
			t.Fatalf("expected canonical encoded path, got %+v", payload.Items[1])
		}
		if payload.Items[2].Path != "テスト2" {
			t.Fatalf("expected canonical unicode path, got %+v", payload.Items[2])
		}
	})

	t.Run("read supports workbook, sheet, and range output", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		workbookResult := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "")
		if workbookResult.exitCode != 0 {
			t.Fatalf("expected workbook read success, stderr=%s", workbookResult.stderr)
		}
		var workbookPayload struct {
			Path string `json:"path"`
			Kind string `json:"kind"`
			Data struct {
				SheetCount int `json:"sheetCount"`
			} `json:"data"`
		}
		if err := json.Unmarshal([]byte(workbookResult.stdout), &workbookPayload); err != nil {
			t.Fatalf("expected workbook JSON: %v", err)
		}
		if workbookPayload.Kind != "workbook" || workbookPayload.Data.SheetCount != 3 {
			t.Fatalf("unexpected workbook payload: %+v", workbookPayload)
		}

		sheetResult := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "Hidden%20Sheet")
		if sheetResult.exitCode != 0 {
			t.Fatalf("expected sheet read success, stderr=%s", sheetResult.stderr)
		}
		var sheetPayload struct {
			Path string `json:"path"`
			Kind string `json:"kind"`
			Data struct {
				Name      string `json:"name"`
				Hidden    bool   `json:"hidden"`
				UsedRange string `json:"usedRange"`
			} `json:"data"`
		}
		if err := json.Unmarshal([]byte(sheetResult.stdout), &sheetPayload); err != nil {
			t.Fatalf("expected sheet JSON: %v", err)
		}
		if !sheetPayload.Data.Hidden || sheetPayload.Data.Name != "Hidden Sheet" {
			t.Fatalf("unexpected sheet payload: %+v", sheetPayload)
		}

		rangeResult := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "Data/C2", "--formula")
		if rangeResult.exitCode != 0 {
			t.Fatalf("expected range read success, stderr=%s", rangeResult.stderr)
		}
		var rangePayload struct {
			Path string `json:"path"`
			Kind string `json:"kind"`
			Data struct {
				Formulas [][]any `json:"formulas"`
			} `json:"data"`
		}
		if err := json.Unmarshal([]byte(rangeResult.stdout), &rangePayload); err != nil {
			t.Fatalf("expected range JSON: %v", err)
		}
		if got := rangePayload.Data.Formulas[0][0]; got != "=SUM(1,2)" {
			t.Fatalf("expected formula, got %#v", got)
		}
	})

	t.Run("read returns date and currency values", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		file, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to open workbook: %v", err)
		}
		dateRaw := mustGetRawCellValue(t, file, "Data", "D2")
		currencyRaw := mustGetRawCellValue(t, file, "Data", "E2")
		if err := file.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}

		result := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "Data/D2:E2", "--value")
		if result.exitCode != 0 {
			t.Fatalf("expected range read success, stderr=%s", result.stderr)
		}

		payload := mustParseRangeValuesPayload(t, result.stdout)
		if payload.Backend == "ole" {
			if got := payload.Data.Values[0][0]; got != "2025-01-02" {
				t.Fatalf("expected formatted date from OLE, got %#v", got)
			}
			if got := payload.Data.Values[0][1]; got != "$1,234.50" {
				t.Fatalf("expected formatted currency from OLE, got %#v", got)
			}
			return
		}

		assertJSONValueMatchesRawNumber(t, payload.Data.Values[0][0], dateRaw)
		assertJSONValueMatchesRawNumber(t, payload.Data.Values[0][1], currencyRaw)
	})

	t.Run("write updates values formulas styles and props", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		valueResult := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data/A2:B2", "--value", `[["Alice",95]]`)
		if valueResult.exitCode != 0 {
			t.Fatalf("expected value write success, stderr=%s", valueResult.stderr)
		}
		formulaResult := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data/C2", "--formula", `[[ "=SUM(3,4)" ]]`)
		if formulaResult.exitCode != 0 {
			t.Fatalf("expected formula write success, stderr=%s", formulaResult.stderr)
		}
		styleResult := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data/A1:A1", "--style", `[[{"font":{"bold":true}}]]`)
		if styleResult.exitCode != 0 {
			t.Fatalf("expected style write success, stderr=%s", styleResult.stderr)
		}
		propsResult := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Hidden%20Sheet", "--props", `{"hidden":false}`)
		if propsResult.exitCode != 0 {
			t.Fatalf("expected props write success, stderr=%s", propsResult.stderr)
		}

		file, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to reopen workbook: %v", err)
		}
		defer file.Close()

		if value, _ := file.GetCellValue("Data", "A2"); value != "Alice" {
			t.Fatalf("expected updated value, got %q", value)
		}
		if formula, _ := file.GetCellFormula("Data", "C2"); formula != "=SUM(3,4)" {
			t.Fatalf("expected updated formula, got %q", formula)
		}
		visible, err := file.GetSheetVisible("Hidden Sheet")
		if err != nil {
			t.Fatalf("failed to get visibility: %v", err)
		}
		if !visible {
			t.Fatal("expected Hidden Sheet to become visible")
		}
	})

	t.Run("write persists date and currency values and styles", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		file, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to open workbook: %v", err)
		}
		dateRaw := mustGetRawCellValue(t, file, "Data", "D2")
		if err := file.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}

		valuePayload := fmt.Sprintf("[[%s,9876.5]]", dateRaw)
		valueResult := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data/D3:E3", "--value", valuePayload)
		if valueResult.exitCode != 0 {
			t.Fatalf("expected date/currency value write success, stderr=%s", valueResult.stderr)
		}
		styleResult := runCLICommand(t, binaryPath, workDir, "write", workbookPath, "Data/D3:E3", "--style", `[[{"numFmt":"yyyy-mm-dd"},{"numFmt":"$#,##0.00"}]]`)
		if styleResult.exitCode != 0 {
			t.Fatalf("expected date/currency style write success, stderr=%s", styleResult.stderr)
		}

		reopened, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to reopen workbook: %v", err)
		}
		defer reopened.Close()

		if got := mustGetRawCellValue(t, reopened, "Data", "D3"); got != dateRaw {
			t.Fatalf("expected raw date serial %q, got %q", dateRaw, got)
		}
		if got := mustGetRawCellValue(t, reopened, "Data", "E3"); got != "9876.5" {
			t.Fatalf("expected raw currency value 9876.5, got %q", got)
		}
		if got := mustGetCustomNumFmt(t, reopened, "Data", "D3"); got != "yyyy-mm-dd" {
			t.Fatalf("expected date numFmt, got %q", got)
		}
		if got := mustGetCustomNumFmt(t, reopened, "Data", "E3"); got != "$#,##0.00" {
			t.Fatalf("expected currency numFmt, got %q", got)
		}

		result := runCLICommand(t, binaryPath, workDir, "read", workbookPath, "Data/D3:E3", "--value")
		if result.exitCode != 0 {
			t.Fatalf("expected readback success, stderr=%s", result.stderr)
		}

		payload := mustParseRangeValuesPayload(t, result.stdout)
		if payload.Backend == "ole" {
			if got := payload.Data.Values[0][0]; got != "2025-01-02" {
				t.Fatalf("expected formatted date from OLE, got %#v", got)
			}
			if got := payload.Data.Values[0][1]; got != "$9,876.50" {
				t.Fatalf("expected formatted currency from OLE, got %#v", got)
			}
			return
		}

		assertJSONValueMatchesRawNumber(t, payload.Data.Values[0][0], dateRaw)
		assertJSONValueMatchesRawNumber(t, payload.Data.Values[0][1], "9876.5")
	})

	t.Run("write reads JSON from stdin when payload is dash", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		result := runCLICommandWithInput(t, binaryPath, workDir, `[["Carol",88]]`, "write", workbookPath, "Data/A2:B2", "--value", "-")
		if result.exitCode != 0 {
			t.Fatalf("expected stdin write success, stderr=%s", result.stderr)
		}

		file, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to reopen workbook: %v", err)
		}
		defer file.Close()

		if value, _ := file.GetCellValue("Data", "A2"); value != "Carol" {
			t.Fatalf("expected updated value from stdin, got %q", value)
		}
		if value, _ := file.GetCellValue("Data", "B2"); value != "88" {
			t.Fatalf("expected updated numeric value from stdin, got %q", value)
		}
	})

	t.Run("add and remove create and delete sheets", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		addResult := runCLICommand(t, binaryPath, workDir, "add", workbookPath, "Sales")
		if addResult.exitCode != 0 {
			t.Fatalf("expected add success, stderr=%s", addResult.stderr)
		}
		dryRunResult := runCLICommand(t, binaryPath, workDir, "remove", workbookPath, "Sales")
		if dryRunResult.exitCode != 0 {
			t.Fatalf("expected dry-run remove success, stderr=%s", dryRunResult.stderr)
		}
		forceResult := runCLICommand(t, binaryPath, workDir, "remove", workbookPath, "Sales", "--force")
		if forceResult.exitCode != 0 {
			t.Fatalf("expected forced remove success, stderr=%s", forceResult.stderr)
		}

		file, err := excelize.OpenFile(workbookPath)
		if err != nil {
			t.Fatalf("failed to reopen workbook: %v", err)
		}
		defer file.Close()
		for _, sheet := range file.GetSheetList() {
			if sheet == "Sales" {
				t.Fatalf("expected Sales to be removed, sheets=%v", file.GetSheetList())
			}
		}
	})

	t.Run("export html writes an artifact path", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		result := runCLICommand(t, binaryPath, workDir, "export", workbookPath, "Data/A1:C2", "--format", "html", "--formula", "--style")
		if result.exitCode != 0 {
			t.Fatalf("expected export success, stderr=%s", result.stderr)
		}
		outputPath := strings.TrimSpace(result.stdout)
		if !filepath.IsAbs(outputPath) {
			t.Fatalf("expected absolute path, got %s", outputPath)
		}
		content, err := os.ReadFile(outputPath)
		if err != nil {
			t.Fatalf("failed to read HTML output: %v", err)
		}
		html := string(content)
		if !strings.Contains(html, "=SUM(1,2)") || !strings.Contains(html, "<table>") {
			t.Fatalf("unexpected HTML output: %s", html)
		}
	})

	t.Run("export png fails on non-Windows runtimes without OLE", func(t *testing.T) {
		t.Parallel()
		if runtime.GOOS == "windows" {
			t.Skip("OLE behavior depends on local Excel runtime")
		}

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)
		result := runCLICommand(t, binaryPath, workDir, "export", workbookPath, "Data", "--format", "png")
		if result.exitCode == 0 {
			t.Fatalf("expected png export failure on non-Windows, stdout=%s", result.stdout)
		}
		if !strings.Contains(result.stderr, "PNG capture") && !strings.Contains(result.stderr, "OLE") {
			t.Fatalf("expected OLE-related error, stderr=%s", result.stderr)
		}
	})

	t.Run("path defaults to workbook root when omitted", func(t *testing.T) {
		t.Parallel()

		workDir := t.TempDir()
		workbookPath := createCLIWorkbookFixture(t, workDir)

		readResult := runCLICommand(t, binaryPath, workDir, "read", workbookPath)
		if readResult.exitCode != 0 {
			t.Fatalf("expected read without path to succeed, stderr=%s", readResult.stderr)
		}
		var readPayload struct {
			Kind string `json:"kind"`
		}
		if err := json.Unmarshal([]byte(readResult.stdout), &readPayload); err != nil {
			t.Fatalf("expected JSON output: %v", err)
		}
		if readPayload.Kind != "workbook" {
			t.Fatalf("expected workbook kind when path omitted, got %s", readPayload.Kind)
		}

		queryResult := runCLICommand(t, binaryPath, workDir, "query", workbookPath)
		if queryResult.exitCode != 0 {
			t.Fatalf("expected query without path to succeed, stderr=%s", queryResult.stderr)
		}
		var queryPayload struct {
			Kind string `json:"kind"`
		}
		if err := json.Unmarshal([]byte(queryResult.stdout), &queryPayload); err != nil {
			t.Fatalf("expected JSON output: %v", err)
		}
		if queryPayload.Kind != "sheetCollection" {
			t.Fatalf("expected sheetCollection kind when path omitted, got %s", queryPayload.Kind)
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
	if _, err := workbook.NewSheet("Hidden Sheet"); err != nil {
		t.Fatalf("failed to create hidden sheet: %v", err)
	}
	if _, err := workbook.NewSheet("テスト2"); err != nil {
		t.Fatalf("failed to create unicode sheet: %v", err)
	}
	if err := workbook.SetSheetVisible("Hidden Sheet", false); err != nil {
		t.Fatalf("failed to hide sheet: %v", err)
	}

	mustSetCellValue(t, workbook, "Data", "A1", "Name")
	mustSetCellValue(t, workbook, "Data", "B1", "Score")
	mustSetCellValue(t, workbook, "Data", "C1", "Calc")
	mustSetCellValue(t, workbook, "Data", "A2", "Bob")
	mustSetCellValue(t, workbook, "Data", "B2", 90)
	mustSetCellFormula(t, workbook, "Data", "C2", "=SUM(1,2)")
	mustSetCellValue(t, workbook, "Data", "D1", "Date")
	mustSetCellValue(t, workbook, "Data", "E1", "Amount")
	mustSetCellValue(t, workbook, "Data", "D2", time.Date(2025, time.January, 2, 0, 0, 0, 0, time.UTC))
	mustSetCellValue(t, workbook, "Data", "E2", 1234.5)
	if err := workbook.SetSheetDimension("Data", "A1:E2"); err != nil {
		t.Fatalf("failed to set Data dimension: %v", err)
	}
	if err := workbook.SetSheetDimension("Hidden Sheet", "A1"); err != nil {
		t.Fatalf("failed to set Hidden Sheet dimension: %v", err)
	}
	if err := workbook.SetSheetDimension("テスト2", "A1:B1"); err != nil {
		t.Fatalf("failed to set テスト2 dimension: %v", err)
	}

	styleID, err := workbook.NewStyle(&excelize.Style{
		Border: []excelize.Border{{Type: "left", Style: 1, Color: "FF0000"}},
	})
	if err != nil {
		t.Fatalf("failed to create style: %v", err)
	}
	if err := workbook.SetCellStyle("Data", "C2", "C2", styleID); err != nil {
		t.Fatalf("failed to apply style: %v", err)
	}
	dateNumFmt := "yyyy-mm-dd"
	dateStyleID, err := workbook.NewStyle(&excelize.Style{CustomNumFmt: &dateNumFmt})
	if err != nil {
		t.Fatalf("failed to create date style: %v", err)
	}
	if err := workbook.SetCellStyle("Data", "D2", "D2", dateStyleID); err != nil {
		t.Fatalf("failed to apply date style: %v", err)
	}
	currencyNumFmt := "$#,##0.00"
	currencyStyleID, err := workbook.NewStyle(&excelize.Style{CustomNumFmt: &currencyNumFmt})
	if err != nil {
		t.Fatalf("failed to create currency style: %v", err)
	}
	if err := workbook.SetCellStyle("Data", "E2", "E2", currencyStyleID); err != nil {
		t.Fatalf("failed to apply currency style: %v", err)
	}

	workbookPath := filepath.Join(dir, "fixture.xlsx")
	if err := workbook.SaveAs(workbookPath); err != nil {
		t.Fatalf("failed to save workbook: %v", err)
	}
	return workbookPath
}

func mustSetCellValue(t *testing.T, workbook *excelize.File, sheet, cell string, value any) {
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

func mustGetRawCellValue(t *testing.T, workbook *excelize.File, sheet, cell string) string {
	t.Helper()

	value, err := workbook.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("failed to get raw cell value for %s: %v", cell, err)
	}
	return value
}

func mustGetCustomNumFmt(t *testing.T, workbook *excelize.File, sheet, cell string) string {
	t.Helper()

	styleID, err := workbook.GetCellStyle(sheet, cell)
	if err != nil {
		t.Fatalf("failed to get style id for %s: %v", cell, err)
	}
	style, err := workbook.GetStyle(styleID)
	if err != nil {
		t.Fatalf("failed to get style for %s: %v", cell, err)
	}
	if style.CustomNumFmt == nil {
		return ""
	}
	return *style.CustomNumFmt
}

func mustParseRangeValuesPayload(t *testing.T, stdout string) struct {
	Path    string `json:"path"`
	Kind    string `json:"kind"`
	Backend string `json:"backend"`
	Data    struct {
		Values [][]any `json:"values"`
	} `json:"data"`
} {
	t.Helper()

	var payload struct {
		Path    string `json:"path"`
		Kind    string `json:"kind"`
		Backend string `json:"backend"`
		Data    struct {
			Values [][]any `json:"values"`
		} `json:"data"`
	}
	if err := json.Unmarshal([]byte(stdout), &payload); err != nil {
		t.Fatalf("expected range values JSON: %v\n%s", err, stdout)
	}
	return payload
}

func assertJSONValueMatchesRawNumber(t *testing.T, got any, raw string) {
	t.Helper()

	want, err := strconv.ParseFloat(raw, 64)
	if err != nil {
		t.Fatalf("failed to parse raw numeric value %q: %v", raw, err)
	}
	number, ok := got.(float64)
	if !ok {
		t.Fatalf("expected JSON number for raw value %q, got %#v", raw, got)
	}
	if number != want {
		t.Fatalf("expected JSON number %v, got %v", want, number)
	}
}

func runCLICommand(t *testing.T, binaryPath, workDir string, args ...string) cliRunResult {
	t.Helper()

	return runCLICommandWithInput(t, binaryPath, workDir, "", args...)
}

func runCLICommandWithInput(t *testing.T, binaryPath, workDir, stdin string, args ...string) cliRunResult {
	t.Helper()

	cmd := exec.Command(binaryPath, args...)
	cmd.Dir = workDir
	if stdin != "" {
		cmd.Stdin = strings.NewReader(stdin)
	}
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

func projectRoot(t *testing.T) string {
	t.Helper()

	root, err := filepath.Abs(filepath.Join("..", ".."))
	if err != nil {
		t.Fatalf("failed to resolve project root: %v", err)
	}
	return root
}
