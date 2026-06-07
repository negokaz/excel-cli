---
name: excel-cli
description: Read, inspect, export, and update Excel workbooks.
---

# excel-cli

Available commands:

- `read`: read workbook, sheet, or range data as JSON
- `query`: enumerate sheets from workbook root `/`
- `write`: write values, formulas, styles, or sheet properties
- `add`: add a sheet
- `remove`: remove a sheet
- `export`: export HTML or PNG artifacts under `.excel-cli/`

## Installation

If you don't have `excel-cli` installed, you can install it with:

```shell
npm install -g @negokaz/excel-cli
```

## Workflow

Read an entire sheet:
```shell
excel-cli export book.xlsx /Sheet1 --format html
```

Search for a value in a sheet:
```shell
$html = excel-cli export book.xlsx /Sheet1 --format html
rg "Alice" $html
```

Read a small range:
```shell
excel-cli read book.xlsx /Sheet1/A1:C3 --value
```

Write values to a range:
```shell
# Write with a JSON argument:
excel-cli write book.xlsx /Sheet1/A1 --value '[[123]]'
# Write with JSON from standard input:
echo '[[123]]' | excel-cli write book.xlsx /Sheet1/A1 --value -
```

Check layout or visual formatting:
```shell
excel-cli export book.xlsx /Sheet1/A1:C20 --format png
```

## Path Rules

- Use canonical paths that begin with `/`.
- Supported initial paths are `/`, `/<sheet>`, `/<sheet>/A1`, and `/<sheet>/A1:C3`.
- Do not use Excel-style paths such as `Sheet1!A1:C3`; the command rejects `!`-separated references.
- Percent-encode spaces and reserved ASCII characters in sheet names, such as `/Hidden%20Sheet` or `/Budget%25`.
- Preserve Unicode characters as-is when possible, and use uppercase cell references in generated paths.

## Notes

- `query` is for workbook structure, not cell-content search. In the current design it enumerates sheets only at `/`.
- `read` returns JSON. Range reads support `--value`, `--formula`, and `--style`.
- `write` accepts JSON payloads that match the shape returned by `read`.
- `write` reads the JSON payload from standard input when the selected update argument is `-`.
- `export` writes artifacts under `.excel-cli/` and prints the absolute output path.
- HTML export is good for broad inspection and grep-based workflows.
- PNG export is good for visual layout checks.
- PNG export may require an environment/backend combination that supports capture.
- Prefer HTML export for broad inspection and grep-based workflows. Prefer `read` for small structured reads. Prefer PNG export for layout checks.
