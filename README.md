# excel-cli

`excel-cli` is a command-line tool for treating Excel workbooks, sheets, and ranges as structured resources.

## Features

- Read workbook, sheet, or range data as JSON
- Read formulas and styles in addition to values
- Write values, formulas, styles, and sheet properties from JSON
- Export HTML artifacts under `.excel-cli/` for broad inspection and visual layout checks

**🪟Windows only:**

- Live edits through OLE automation when Excel is available
- Export PNG captures under `.excel-cli/` through OLE automation

## Installation

Requires Node.js 20 or later.

```sh
npm install -g @negokaz/excel-cli
```

## Installing the Agent Skill

This repository also includes an agent skill at [`skills/excel-cli/`](./skills/excel-cli/).

You can install the skill with:

```sh
gh skill install negokaz/excel-cli excel-cli
```

## Command Summary

```text
excel-cli read <file> <path> [--value | --formula | --style]
excel-cli query <file> <path>
excel-cli write <file> <path> (--value <json|-> | --formula <json|-> | --style <json|-> | --props <json|->)
excel-cli add <file> <path>
excel-cli remove <file> <path> [--force]
excel-cli export <file> <path> --format <html|png> [--formula] [--style]
```

## Paths

Commands address workbook resources through a canonical `<file> <path>` pair.

Supported initial paths:

- `/`
- `/Sheet1`
- `/Sheet1/A1`
- `/Sheet1/A1:C3`

Path rules:

- paths must begin with `/`
- sheet names use canonical path segments: Unicode characters are preserved, while ASCII characters that require escaping remain percent-encoded
- range output is canonicalized to uppercase cell references
- Excel-style references such as `Sheet1!A1:C3` are rejected

## Examples

Enumerate sheets:

```sh
excel-cli query book.xlsx /
```

```json
{
  "path": "/",
  "kind": "sheetCollection",
  "backend": "excelize",
  "items": [
    { "path": "/Data", "kind": "sheet", "name": "Data" },
    { "path": "/Hidden%20Sheet", "kind": "sheet", "name": "Hidden Sheet" }
  ]
}
```

Read workbook, sheet, and range resources:

```sh
excel-cli read book.xlsx /
excel-cli read book.xlsx /Data
excel-cli read book.xlsx /Data/A1:C2 --formula
excel-cli read book.xlsx /Data/A1:C2 --style
```

Write values, formulas, styles, and sheet properties:

```sh
excel-cli write book.xlsx /Data/A2:B2 --value '[["Alice",95]]'
excel-cli write book.xlsx /Data/C2 --formula '[["=SUM(3,4)"]]'
excel-cli write book.xlsx /Data/A1:B1 --style '[[{"font":{"bold":true}}, null]]'
excel-cli write book.xlsx /Hidden%20Sheet --props '{"hidden":false}'
echo '[[123]]' | excel-cli write book.xlsx /Data/A1 --value -
```

Create and remove worksheets:

```sh
excel-cli add book.xlsx /Sales
excel-cli remove book.xlsx /Sales
excel-cli remove book.xlsx /Sales --force
```

`remove` is a dry-run by default. It validates that the sheet can be removed and returns JSON with `wouldRemove: true`. The workbook is only changed when `--force` is provided.

Export derived artifacts:

```sh
excel-cli export book.xlsx /Data --format html --formula --style
excel-cli export book.xlsx /Data/A1:C10 --format png
```

HTML and PNG files are created under `.excel-cli/`.

## Notes

- `write --value`, `write --formula`, and `read --value`, `read --formula` are designed to round-trip through the same 2-dimensional JSON shape
- `write --style` accepts a 2-dimensional array of style objects or `null`
- `write --props` currently supports only worksheet `hidden`
- `write` reads the JSON payload from standard input when the selected update argument is `-`
- `remove` fails if the target sheet does not exist or if it is the workbook's only worksheet

The design notes in [docs/](docs/index.md) are the primary reference for command contracts and migration intent.

## Supported Platforms

| Platform | Architecture |
|----------|--------------|
| Windows  | x64, arm64   |
| macOS    | x64, arm64   |
| Linux    | x64, arm64   |

## License

MIT
