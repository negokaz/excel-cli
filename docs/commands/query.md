---
tags:
  - excel-cli
  - design
  - query
---

# query

This note describes the behavior that callers can rely on when using `excel-cli query`. It defines collection semantics, JSON output shape, and failure conditions.

## Command Form

`excel-cli query <file> <path>`

- `<file>` is the workbook to inspect.
- `<path>` is the collection path to enumerate.

The command requires both `<file>` and `<path>`.

## Supported Paths

In the initial version, `query` formally supports only the workbook root:

- `/`

The meaning of `query /` is "enumerate the workbook's direct sheet collection".

`query` is not a recursive search command in the initial version. It does not search descendants or perform pattern matching.

## Successful Behavior

On success, `query` writes formatted JSON to standard output and exits successfully.

The response is a wrapper object, not a bare array.

All successful responses include:

- `path`, always returned in canonical form
- `kind`, `sheetCollection` for the initial version
- `backend`, the backend that opened the workbook
- `items`, the enumerated items

Example:

```json
{
  "path": "/",
  "kind": "sheetCollection",
  "backend": "ole",
  "items": [
    {
      "path": "/Sheet1",
      "kind": "sheet",
      "name": "Sheet1"
    },
    {
      "path": "/Sales",
      "kind": "sheet",
      "name": "Sales"
    }
  ]
}
```

The item order follows workbook sheet order.

## Failure Conditions

`query` fails when:

- the workbook path cannot be resolved or opened
- `<path>` does not follow supported path syntax
- `<path>` is valid but not a supported collection path in the initial version
- worksheet metadata cannot be enumerated
- required arguments are missing
- the JSON response cannot be encoded

Callers can expect errors to distinguish at least these cases:

- invalid path syntax
- unsupported path kind

## Scope Notes

This note defines the initial external contract of `query`. It does not define future collection paths such as tables, charts, or named ranges.

## Related Notes

- [[core-concept]]
- [[read]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[read]: read "Behavior of the read Command"
