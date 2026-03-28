package cli

import (
	"fmt"
	"os"
)

const usage = `excel-cli - Excel file reader

Usage:
  excel-cli list <file>
  excel-cli read <file> <sheet> [options]

Commands:
  list    List all sheets in the Excel file
  read    Read sheet content and save as HTML

Options for read:
  --formula   Show formulas instead of values
  --style     Include cell style information

Output:
  The read command writes the sheet content to:
    .excel-cli/sheet-<timestamp>.html
  and prints the absolute path to stdout.`

func Run(args []string) error {
	if len(args) == 0 {
		fmt.Fprintln(os.Stderr, usage)
		return nil
	}
	switch args[0] {
	case "list":
		return runList(args[1:])
	case "read":
		return runRead(args[1:])
	case "help", "--help", "-h":
		fmt.Fprintln(os.Stderr, usage)
		return nil
	default:
		fmt.Fprintln(os.Stderr, usage)
		return fmt.Errorf("unknown command: %s", args[0])
	}
}
