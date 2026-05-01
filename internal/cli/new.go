package cli

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/negokaz/excel-cli/internal/excel"
)

func runNew(args []string) error {
	if len(args) == 0 {
		return fmt.Errorf("usage: excel-cli new <file>")
	}

	filePath := args[0]
	absPath, err := filepath.Abs(filePath)
	if err != nil {
		return fmt.Errorf("failed to resolve path: %w", err)
	}

	if err := excel.NewFile(absPath); err != nil {
		return err
	}

	fmt.Fprintln(os.Stdout, absPath)
	return nil
}
