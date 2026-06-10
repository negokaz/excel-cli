---
tags:
  - excel-cli
  - design
  - add
---

# add

This note describes the behavior that callers can rely on when using `excel-cli add`. It defines the creation contract, supported targets, and failure conditions.

## Command Form

`excel-cli add <file> <path>`

- `<file>` is the workbook to update.
- `<path>` is the canonical path to create.

The command requires both `<file>` and `<path>`.

## Supported Targets

In the initial version, `add` supports only worksheet creation.

The supported form is:

- `<sheet>`

Example:

```text
excel-cli add book.xlsx Sales
```

## Successful Behavior

On success, `add` creates the requested sheet at the end of the workbook, saves the workbook in place, writes formatted JSON to standard output, and exits successfully.

Example:

```json
{
  "path": "Sales",
  "kind": "sheet",
  "action": "add"
}
```

Observable behavior includes:

- the path must not already exist
- the new sheet is appended to the workbook's sheet order
- no additional properties are accepted at creation time

## Failure Conditions

`add` fails when:

- the workbook path cannot be resolved or opened
- `<path>` does not follow supported path syntax
- `<path>` is not a supported creatable path in the initial version
- the target sheet already exists
- required arguments are missing
- the workbook cannot be saved

Callers can expect errors to distinguish at least these cases:

- invalid path syntax
- unsupported path kind

## Scope Notes

This note defines the initial external contract of `add`. It does not define creation of tables, charts, named ranges, or workbook-level resources.

## Related Notes

- [[core-concept]]
- [[write]]
- [[remove]]
- [[query]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[write]: write "Behavior of the write Command"
[remove]: remove "Behavior of the remove Command"
[query]: query "Behavior of the query Command"
