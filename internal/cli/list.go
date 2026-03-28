package cli

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runList(args []string) error {
	if len(args) == 0 {
		return fmt.Errorf("usage: excel-cli list <file>")
	}

	filePath := args[0]
	absPath, err := filepath.Abs(filePath)
	if err != nil {
		return fmt.Errorf("failed to resolve path: %w", err)
	}

	workbook, release, err := excel.OpenFile(absPath)
	if err != nil {
		return fmt.Errorf("failed to open file: %w", err)
	}
	defer release()

	sheets, err := workbook.GetSheets()
	if err != nil {
		return fmt.Errorf("failed to get sheets: %w", err)
	}

	type SheetInfo struct {
		Name      string `json:"name"`
		UsedRange string `json:"usedRange"`
	}
	type Response struct {
		Backend string      `json:"backend"`
		Sheets  []SheetInfo `json:"sheets"`
	}

	sheetInfos := make([]SheetInfo, 0, len(sheets))
	for _, sheet := range sheets {
		defer sheet.Release()
		name, err := sheet.Name()
		if err != nil {
			return err
		}
		usedRange, _ := sheet.GetDimention()
		sheetInfos = append(sheetInfos, SheetInfo{Name: name, UsedRange: usedRange})
	}

	out, err := json.MarshalIndent(Response{Backend: workbook.GetBackendName(), Sheets: sheetInfos}, "", "  ")
	if err != nil {
		return err
	}
	fmt.Fprintln(os.Stdout, string(out))
	return nil
}
