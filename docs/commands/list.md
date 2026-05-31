---
tags:
  - excel-cli
  - design
  - list
---

# list

This note describes the behavior that callers can rely on when using `excel-cli list`. It focuses on the command-line contract and the JSON output, not on how the tool chooses an internal workbook backend.

## Command Form

`excel-cli list <file>`

- `<file>` is the workbook to inspect.

The command requires `<file>`. If the argument is missing, the command fails with a usage error.

## Input Semantics

### Workbook Selection

`list` opens the specified workbook and enumerates its worksheets.

The output order follows the workbook's sheet order, so callers can treat the returned array as an ordered sheet listing rather than an unordered set.

### Output Shape

On success, `list` writes a JSON object to standard output.

The top-level object contains:

- `backend`, the backend name that opened the workbook, such as `ole` or `excelize`
- `sheets`, an array of sheet metadata objects

Each sheet metadata object contains:

- `name`, the worksheet name
- `usedRange`, the backend-reported used range for that sheet

## Successful Behavior

On success, `list` exits successfully after printing valid JSON to standard output.

Observable characteristics of the output include:

- one entry per worksheet in the workbook
- the backend name used for the workbook session
- a used-range string for each sheet, such as `A1:C2`
- formatted JSON with indentation rather than a single compact line

Unlike `read` and `capture`, this command does not create a generated artifact under `.excel-cli/`.

## Failure Conditions

`list` fails when:

- the workbook path cannot be resolved or opened
- worksheet metadata cannot be enumerated
- a required argument is missing
- the JSON response cannot be encoded

When the command fails, it exits with an error instead of printing a JSON payload that callers should consume.

## Scope Notes

This note defines the external contract of `list`: what input it requires and what JSON structure it emits. It does not define backend-selection rules or behavior for unspecified trailing arguments.

## Related Notes

- [[core-concept]]
- [[read]]
- [[capture]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[read]: read "Behavior of the read Command"
[capture]: capture "Behavior of the capture Command"