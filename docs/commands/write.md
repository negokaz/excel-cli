---
tags:
  - excel-cli
  - design
  - write
---

# write

This note describes the behavior that callers can rely on when using `excel-cli write`. It focuses on the observable command contract rather than internal workbook handling.

## Command Form

`excel-cli write <file> <sheet> <range> <values> [--newsheet]`

- `<file>` is the workbook to update.
- `<sheet>` is the target sheet name.
- `<range>` is the target Excel range.
- `<values>` is a JSON 2-dimensional array.
- `--newsheet` creates the target sheet before writing.

The command requires four positional arguments. It does not create a new workbook, so the workbook file must already exist.

## Input Contract

### Range

The range is expressed separately from the sheet name. Inputs such as `Sheet1!A1:C3` are rejected. Valid inputs are standard Excel-style cell references and rectangular ranges such as `A1`, `A1:C2`, or `$A$1:$C$2`.

### Values

`<values>` must be a JSON 2-dimensional array.

- the outer array represents rows
- each inner array represents cells in that row
- the shape must match the target range exactly
- a one-dimensional JSON array is rejected
- rows with the wrong column count are rejected

Strings that begin with `=` are treated as formulas rather than plain text values.

## Sheet Creation and Selection

Without `--newsheet`, the target sheet must already exist.

With `--newsheet`, `write` creates the target sheet and then writes the provided values. This mode is exclusive with existing sheets: if the sheet already exists, the command fails instead of reusing it.

## Successful Behavior

On success, `write` updates the workbook in place and exits successfully.

Observable write semantics include:

- values in the target range are overwritten
- formula strings that begin with `=` are written as formulas
- the input shape must match the target range exactly

Unlike `new` and `read`, this command does not print an output path on success.

## Failure Conditions

`write` fails when:

- the workbook path cannot be resolved or opened
- the target sheet does not exist and `--newsheet` is not set
- `--newsheet` is set but the target sheet already exists
- the range is invalid or includes a sheet name
- `<values>` is not valid JSON
- `<values>` is not a JSON 2-dimensional array
- the row count does not match the range height
- any row's column count does not match the range width
- required arguments are missing
- extra arguments are provided

When the command fails, it exits with an error. Success is signaled by the updated workbook file and a successful process exit status, not by a generated artifact path.

## Scope Notes

This note defines the external contract of `write`: what the caller provides, what the command updates, and which inputs are rejected. It does not define backend-specific type coercion, internal write strategy, or spreadsheet-application behavior beyond what is observable from the command itself.

## Related Notes

- [[core-concept]]
- [[read]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[read]: read "Behavior of the read Command"
