package cli

import (
	"flag"
	"fmt"
	"path/filepath"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runRemove(args []string) error {
	if len(args) < 1 {
		return fmt.Errorf("usage: excel-cli remove <file> [<path>] [--force]")
	}

	fs := flag.NewFlagSet("remove", flag.ContinueOnError)
	force := fs.Bool("force", false, "Delete the target instead of dry-run validation")

	filePath := args[0]
	rawPath, remaining := extractPathArg(args[1:])
	if err := fs.Parse(remaining); err != nil {
		return err
	}
	if len(fs.Args()) != 0 {
		return fmt.Errorf("usage: excel-cli remove <file> [<path>] [--force]")
	}

	target, err := parseTargetPath(rawPath)
	if err != nil {
		return err
	}
	if target.Kind != pathKindSheet {
		return fmt.Errorf("unsupported path kind: %s", rawPath)
	}

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
	for _, sheet := range sheets {
		defer sheet.Release()
	}
	if len(sheets) <= 1 {
		return fmt.Errorf("cannot remove the only worksheet")
	}

	worksheet, err := workbook.FindSheet(target.SheetName)
	if err != nil {
		return fmt.Errorf("sheet not found: %s", target.SheetName)
	}
	worksheet.Release()

	if !*force {
		return writeJSON(struct {
			Path        string `json:"path"`
			Kind        string `json:"kind"`
			Action      string `json:"action"`
			WouldRemove bool   `json:"wouldRemove"`
		}{
			Path:        target.Canonical(),
			Kind:        "sheet",
			Action:      "remove",
			WouldRemove: true,
		})
	}

	if err := workbook.DeleteSheet(target.SheetName); err != nil {
		return fmt.Errorf("failed to delete sheet: %w", err)
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
		Action: "remove",
	})
}
