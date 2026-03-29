package cli

import (
	"fmt"
	"os"
)

const usage = `excel-cli - Excel file tool

Usage:
  excel-cli list <file>
  excel-cli read <file> <sheet> [options]
  excel-cli write <file> <sheet> <range> <values> [--newsheet]

Commands:
  list    List all sheets in the Excel file
  read    Read sheet content and save as HTML
  write   Write values to a sheet in the Excel file

Options for read:
  --formula   Show formulas instead of values
  --style     Include cell style information

Options for write:
  --newsheet  Create the sheet if it does not exist (error if it already exists)

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
	case "write":
		return runWrite(args[1:])
	case "help", "--help", "-h":
		fmt.Fprintln(os.Stderr, usage)
		return nil
	default:
		fmt.Fprintln(os.Stderr, usage)
		return fmt.Errorf("unknown command: %s", args[0])
	}
}
