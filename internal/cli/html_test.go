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

	table, _, err := createHTMLTableOfValues(ws, 1, 1, 2, 2)

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

	tableHTML, css, err := createHTMLTableOfValuesWithStyle(ws, 1, 1, 2, 1)

	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}
	// Border CSS rule should be defined exactly once (deduplication)
	if strings.Count(css, ".b1 {") != 1 {
		t.Fatalf("expected exactly one b1 CSS rule, got css: %s", css)
	}
	// Both cells should reference the b1 class
	if strings.Count(tableHTML, `class="b1`) != 2 {
		t.Fatalf("expected both cells to use b1 class, got: %s", tableHTML)
	}
}

func TestGenerateStylesheetSortsStableIDs(t *testing.T) {
	t.Parallel()

	registry := newStyleRegistry()
	registry.borderStyles["b2"] = "border-right: 1px solid"
	registry.borderStyles["b1"] = "border-left: 1px solid"

	css := registry.generateStylesheet()

	firstIndex := strings.Index(css, ".b1 {")
	secondIndex := strings.Index(css, ".b2 {")
	if firstIndex < 0 || secondIndex < 0 || firstIndex > secondIndex {
		t.Fatalf("expected sorted CSS rules, got %s", css)
	}
}

func TestCalculateHashIsStable(t *testing.T) {
	t.Parallel()

	hashA := calculateHash("border-left: 1px solid #ff0000")
	hashB := calculateHash("border-left: 1px solid #ff0000")

	if hashA == "" || hashA != hashB {
		t.Fatalf("expected stable hash, got %s and %s", hashA, hashB)
	}
	if calculateHash("") != "" {
		t.Fatal("expected empty string for empty input")
	}
}

func TestCreateHTMLTableOfValuesWithMergedCells(t *testing.T) {
	t.Parallel()

	ws := openHTMLTestWorksheet(t, func(file *excelize.File) {
		if err := file.SetCellValue("Sheet1", "A1", "merged"); err != nil {
			t.Fatalf("failed to set A1: %v", err)
		}
		if err := file.MergeCell("Sheet1", "A1", "B2"); err != nil {
			t.Fatalf("failed to merge A1:B2: %v", err)
		}
		if err := file.SetCellValue("Sheet1", "C1", "other"); err != nil {
			t.Fatalf("failed to set C1: %v", err)
		}
	})

	table, _, err := createHTMLTableOfValues(ws, 1, 1, 3, 2)

	if err != nil {
		t.Fatalf("expected no error, got %v", err)
	}
	if !strings.Contains(table, `colspan="2"`) {
		t.Fatalf("expected colspan=2, got %s", table)
	}
	if !strings.Contains(table, `rowspan="2"`) {
		t.Fatalf("expected rowspan=2, got %s", table)
	}
	// Covered cells (B1, A2, B2) must not appear as separate <td> elements
	// The table should have fewer <td> elements than cells in the range
	tdCount := strings.Count(table, "<td")
	// 3 cols x 2 rows = 6 cells, but B1/A2/B2 are covered → 3 <td> elements
	if tdCount != 3 {
		t.Fatalf("expected 3 <td> elements (merged region + C1 + C2), got %d in:\n%s", tdCount, table)
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
