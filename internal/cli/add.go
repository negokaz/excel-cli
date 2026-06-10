package cli

import (
	"fmt"
	"path/filepath"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runAdd(args []string) error {
	if len(args) < 1 || len(args) > 2 {
		return fmt.Errorf("usage: excel-cli add <file> [<path>]")
	}

	rawPath := ""
	if len(args) == 2 {
		rawPath = args[1]
	}
	target, err := parseTargetPath(rawPath)
	if err != nil {
		return err
	}
	if target.Kind != pathKindSheet {
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

	if existing, err := workbook.FindSheet(target.SheetName); err == nil {
		existing.Release()
		return fmt.Errorf("sheet already exists: %s", target.SheetName)
	}
	if err := workbook.CreateNewSheet(target.SheetName); err != nil {
		return fmt.Errorf("failed to create sheet: %w", err)
	}
	if err := workbook.Save(); err != nil {
		return fmt.Errorf("failed to save workbook: %w", err)
	}

	return writeJSON(struct {
		Path   string `json:"path"`
		Kind   string `json:"kind"`
		Action string `json:"action"`
	}{
		Path:   target.Canonical(),
		Kind:   "sheet",
		Action: "add",
	})
}
