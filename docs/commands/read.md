---
tags:
  - excel-cli
  - design
  - read
---

# read

This note describes the behavior that callers can rely on when using `excel-cli read`. It defines the command contract, supported path kinds, output shape, and failure conditions.

## Command Form

`excel-cli read <file> <path> [--value | --formula | --style]`

- `<file>` is the workbook to inspect.
- `<path>` is the canonical target path inside the workbook.
- `--value` returns displayed values for a range target.
- `--formula` returns formulas where present, and normal values otherwise, for a range target.
- `--style` returns style data for a range target.

The command requires `<file>` and `<path>`.

The options `--value`, `--formula`, and `--style` are mutually exclusive. If none is specified, range targets default to `--value`.

## Supported Paths

Initial supported path kinds are:

- `/` for the workbook root
- `/<sheet>` for a worksheet
- `/<sheet>/A1` for a single cell, treated as a `range`
- `/<sheet>/A1:C3` for a rectangular range

Path rules:

- the path must begin with `/`
- sheet names use canonical path segments: Unicode characters are preserved, while ASCII characters that require escaping remain percent-encoded
- cell columns are uppercase in canonical output
- Excel-style references such as `Sheet1!A1:C3` are not accepted

## Successful Behavior

On success, `read` writes formatted JSON to standard output and exits successfully.

All successful responses include:

- `path`, always returned in canonical form
- `kind`, one of `workbook`, `sheet`, or `range`
- `backend`, the backend that opened the workbook
- `data`, the payload for the requested target

### Workbook Output

For `/`, `read` returns minimal workbook metadata.

Example:

```json
{
  "path": "/",
  "kind": "workbook",
  "backend": "ole",
  "data": {
    "sheetCount": 3
  }
}
```

### Sheet Output

For `/<sheet>`, `read` returns minimal sheet metadata.

- `name`
- `hidden`
- `usedRange`

Example:

```json
{
  "path": "/Sheet1",
  "kind": "sheet",
  "backend": "ole",
  "data": {
    "name": "Sheet1",
    "hidden": false,
    "usedRange": "A1:C20"
  }
}
```

### Range Output

For range targets, `read` returns data that is compatible with the corresponding `write` input shape.

- `--value` returns a JSON 2-dimensional array of values
- `--formula` returns a JSON 2-dimensional array where formula cells contain formula strings and non-formula cells contain normal values
- `--style` returns a JSON 2-dimensional array of style objects or `null`

On the OLE backend, cells whose COM value is typed as `Date` or `Currency` are returned using each cell's displayed text.

Example with `--value`:

```json
{
  "path": "/Sheet1/A1:B2",
  "kind": "range",
  "backend": "ole",
  "data": {
    "values": [
      ["Name", "Score"],
      ["Alice", 95]
    ]
  }
}
```

Example with `--formula`:

```json
{
  "path": "/Sheet1/A1:B2",
  "kind": "range",
  "backend": "ole",
  "data": {
    "formulas": [
      ["Name", "=SUM(B3:B10)"],
      ["Alice", 95]
    ]
  }
}
```

Example with `--style`:

```json
{
  "path": "/Sheet1/A1:B1",
  "kind": "range",
  "backend": "ole",
  "data": {
    "styles": [
      [
        { "font": { "bold": true } },
        null
      ]
    ]
  }
}
```

## Failure Conditions

`read` fails when:

- the workbook path cannot be resolved or opened
- `<path>` does not follow supported path syntax
- the path kind is not supported in the initial version
- the target sheet does not exist
- the requested range cannot be read
- more than one of `--value`, `--formula`, or `--style` is provided
- required arguments are missing
- the JSON response cannot be encoded

Callers can expect errors to distinguish at least these cases:

- invalid path syntax
- sheet not found
- unsupported path kind

## Scope Notes

This note defines the caller-visible contract of `read`. It does not define backend-specific evaluation differences in formulas, formatting, or workbook metadata beyond what is surfaced in the JSON payload.

## Related Notes

- [[core-concept]]
- [[query]]
- [[write]]
- [[export]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[query]: query "Behavior of the query Command"
[write]: write "Behavior of the write Command"
[export]: export "Behavior of the export Command"
