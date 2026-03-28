package cli

import (
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/negokaz/excel-cli/internal/excel"
	"github.com/xuri/excelize/v2"
)

func TestBuildHTMLPageEscapesMetadataAndKeepsTableHTML(t *testing.T) {
	t.Parallel()

	params := HTMLPageParams{
		FilePath:    `C:\tmp\<book>.xlsx`,
		SheetName:   `<Summary>`,
		UsedRange:   "A1:B2",
		Backend:     "excelize",
		GeneratedAt: time.Date(2026, 3, 28, 7, 30, 0, 0, time.FixedZone("JST", 9*60*60)),
		TableHTML:   "<table><tr><td>raw</td></tr></table>",
	}

	page := buildHTMLPage(params)

	if !strings.Contains(page, "&lt;Summary&gt; - C:\\tmp\\&lt;book&gt;.xlsx") {
		t.Fatalf("expected escaped title, got %s", page)
	}
	if !strings.Contains(page, "Generated: 2026-03-27T22:30:00Z") {
		t.Fatalf("expected UTC timestamp, got %s", page)
	}
	if !strings.Contains(page, params.TableHTML) {
		t.Fatalf("expected raw table HTML to be embedded, got %s", page)
	}
}

func TestCreateHTMLTableOfValuesIncludesHeadersAndLineBreaks(t *testing.T) {
	t.Parallel()

	ws := openHTMLTestWorksheet(t, func(file *excelize.File) {
		if err := file.SetCellValue("Sheet1", "A1", "Header"); err != nil {
			t.Fatalf("failed to set A1: %v", err)
		}
		if err := file.SetCellValue("Sheet1", "B2", "Line 1\nLine 2"); err != nil {
			t.Fatalf("failed to set B2: %v", err)
		}
	})

	table, err := createHTMLTableOfValues(ws, 1, 1, 2, 2)

	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}
	if !strings.Contains(table, "<th>A</th><th>B</th>") {
		t.Fatalf("expected column headers, got %s", table)
	}
	if !strings.Contains(table, "<th>1</th>") || !strings.Contains(table, "<th>2</th>") {
		t.Fatalf("expected row headers, got %s", table)
	}
	if !strings.Contains(table, "Line 1<br>Line 2") {
		t.Fatalf("expected line break conversion, got %s", table)
	}
}

func TestCreateHTMLTableWithStyleDeduplicatesStyleDefinitions(t *testing.T) {
	t.Parallel()

	ws := openHTMLTestWorksheet(t, func(file *excelize.File) {
		if err := file.SetCellValue("Sheet1", "A1", "first"); err != nil {
			t.Fatalf("failed to set A1: %v", err)
		}
		if err := file.SetCellValue("Sheet1", "B1", "second"); err != nil {
			t.Fatalf("failed to set B1: %v", err)
		}
		styleID, err := file.NewStyle(&excelize.Style{
			Border: []excelize.Border{{Type: "left", Style: 1, Color: "FF0000"}},
		})
		if err != nil {
			t.Fatalf("failed to create style: %v", err)
		}
		if err := file.SetCellStyle("Sheet1", "A1", "B1", styleID); err != nil {
			t.Fatalf("failed to set style: %v", err)
		}
	})

	table, err := createHTMLTableOfValuesWithStyle(ws, 1, 1, 2, 1)

	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}
	if strings.Count(table, `id="b1"`) != 1 {
		t.Fatalf("expected one border definition, got %s", table)
	}
	if strings.Count(table, `style-ref="b1 f1 l1"`) != 2 {
		t.Fatalf("expected both cells to reuse the same style-ref set, got %s", table)
	}
}

func TestGenerateStyleDefinitionsSortsStableIDs(t *testing.T) {
	t.Parallel()

	registry := newStyleRegistry()
	registry.borderStyles["b2"] = "second"
	registry.borderStyles["b1"] = "first"

	definitions := registry.generateStyleDefinitions()

	firstIndex := strings.Index(definitions, `id="b1"`)
	secondIndex := strings.Index(definitions, `id="b2"`)
	if firstIndex < 0 || secondIndex < 0 || firstIndex > secondIndex {
		t.Fatalf("expected sorted style definitions, got %s", definitions)
	}
}

func TestConvertToYAMLFlowAndCalculateYAMLHashAreStable(t *testing.T) {
	t.Parallel()

	value := excel.Border{Type: excel.BorderTypeLeft, Style: excel.BorderStyleContinuous, Color: "#FF0000"}

	yamlFlow := convertToYAMLFlow(value)
	hashA := calculateYAMLHash(yamlFlow)
	hashB := calculateYAMLHash(yamlFlow)

	if yamlFlow == "" {
		t.Fatal("expected YAML flow output")
	}
	if strings.Contains(yamlFlow, "\"") {
		t.Fatalf("expected YAML flow without quotes, got %s", yamlFlow)
	}
	if hashA == "" || hashA != hashB {
		t.Fatalf("expected stable hash, got %s and %s", hashA, hashB)
	}
}

func openHTMLTestWorksheet(t *testing.T, configure func(file *excelize.File)) excel.Worksheet {
	t.Helper()

	workbookPath := filepath.Join(t.TempDir(), "html-test.xlsx")
	file := excelize.NewFile()
	configure(file)
	if err := file.SaveAs(workbookPath); err != nil {
		t.Fatalf("failed to save workbook: %v", err)
	}
	if err := file.Close(); err != nil {
		t.Fatalf("failed to close workbook: %v", err)
	}

	workbook, release, err := excel.OpenFile(workbookPath)
	if err != nil {
		t.Fatalf("failed to open workbook: %v", err)
	}
	t.Cleanup(release)

	ws, err := workbook.FindSheet("Sheet1")
	if err != nil {
		t.Fatalf("failed to find Sheet1: %v", err)
	}
	return ws
}
