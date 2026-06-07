package cli

import (
	"fmt"
	"path/filepath"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runQuery(args []string) error {
	if len(args) != 2 {
		return fmt.Errorf("usage: excel-cli query <file> <path>")
	}

	target, err := parseTargetPath(args[1])
	if err != nil {
		return err
	}
	if target.Kind != pathKindWorkbook {
		return fmt.Errorf("unsupported path kind: %s", args[1])
	}

	absPath, err := filepath.Abs(args[0])
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
	items := make([]struct {
		Path string `json:"path"`
		Kind string `json:"kind"`
		Name string `json:"name"`
	}, 0, len(sheets))
	for _, sheet := range sheets {
		name, err := sheet.Name()
		sheet.Release()
		if err != nil {
			return fmt.Errorf("failed to get sheet metadata: %w", err)
		}
		items = append(items, struct {
			Path string `json:"path"`
			Kind string `json:"kind"`
			Name string `json:"name"`
		}{
			Path: canonicalSheetPath(name),
			Kind: "sheet",
			Name: name,
		})
	}

	return writeJSON(struct {
		Path    string `json:"path"`
		Kind    string `json:"kind"`
		Backend string `json:"backend"`
		Items   any    `json:"items"`
	}{
		Path:    "/",
		Kind:    "sheetCollection",
		Backend: workbook.GetBackendName(),
		Items:   items,
	})
}
