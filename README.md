# excel-cli

A CLI tool for reading and writing Excel files. Supports Windows, macOS, and Linux.

On Windows, it uses Excel via OLE automation when Excel is installed, providing accurate results for complex spreadsheets. On other platforms (or when Excel is unavailable), it falls back to the [excelize](https://github.com/xuri/excelize) library.

## Installation

Requires Node.js >= 20.

```sh
npm install -g @negokaz/excel-cli
```

## Usage

```
excel-cli list <file>
excel-cli read <file> <sheet> [options]
excel-cli write <file> <sheet> <range> <values> [--newsheet]
```

### `list` — List sheets

Lists all sheets in an Excel file. Outputs JSON.

```sh
excel-cli list book.xlsx
```

```json
{
  "backend": "ole",
  "sheets": [
    { "name": "Sheet1", "usedRange": "A1:D10" },
    { "name": "Sheet2", "usedRange": "A1:B3" }
  ]
}
```

### `read` — Read a sheet

Reads sheet content and saves it as an HTML file. Prints the absolute path of the output file to stdout.

```sh
excel-cli read book.xlsx Sheet1
```

The output is written to:
```
.excel-cli/sheet-<timestamp>.html
```

**Options:**

| Option      | Description                            |
|-------------|----------------------------------------|
| `--formula` | Show formulas instead of cell values   |
| `--style`   | Include cell style information         |

```sh
# Show formulas with style information
excel-cli read book.xlsx Sheet1 --formula --style
```

### `write` — Write values to a sheet

Writes values to a cell range. The `<values>` argument must be a JSON 2-dimensional array.

```sh
excel-cli write book.xlsx Sheet1 A1:C2 '[["Name","Age","City"],["Alice",30,"Tokyo"]]'
```

**Options:**

| Option        | Description                                                           |
|---------------|-----------------------------------------------------------------------|
| `--newsheet`  | Create the sheet if it does not exist (error if it already exists)   |

```sh
# Write to a new sheet
excel-cli write book.xlsx NewSheet A1 '[["Hello"]]' --newsheet
```

## Supported Platforms

| Platform | Architecture |
|----------|-------------|
| Windows  | x64, arm64  |
| macOS    | x64, arm64  |
| Linux    | x64, arm64  |

## License

MIT
