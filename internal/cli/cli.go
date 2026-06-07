package cli

import (
	"fmt"
	"os"
)

const usage = `excel-cli - Excel file tool

Usage:
  excel-cli read <file> <path> [--value | --formula | --style]
  excel-cli query <file> <path>
  excel-cli write <file> <path> (--value <json|-> | --formula <json|-> | --style <json|-> | --props <json|->)
  excel-cli add <file> <path>
  excel-cli remove <file> <path> [--force]
  excel-cli export <file> <path> --format <html|png> [options]

Commands:
  read     Read workbook, sheet, or range data as JSON
  query    Enumerate collection resources as JSON
  write    Update workbook resources from JSON payloads
  add      Create workbook resources
  remove   Validate or delete workbook resources
  export   Render HTML or PNG artifacts under .excel-cli/

Output:
  read/query/write/add/remove return JSON on success.
  export writes an artifact under .excel-cli/ and prints its absolute path.`

func Run(args []string) error {
	if len(args) == 0 {
		fmt.Fprintln(os.Stderr, usage)
		return nil
	}
	switch args[0] {
	case "read":
		return runRead(args[1:])
	case "query":
		return runQuery(args[1:])
	case "write":
		return runWrite(args[1:])
	case "add":
		return runAdd(args[1:])
	case "remove":
		return runRemove(args[1:])
	case "export":
		return runExport(args[1:])
	case "help", "--help", "-h":
		fmt.Fprintln(os.Stderr, usage)
		return nil
	default:
		fmt.Fprintln(os.Stderr, usage)
		return fmt.Errorf("unknown command: %s", args[0])
	}
}
