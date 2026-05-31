---
tags:
  - excel-cli
  - design
  - capture
---

# capture

This note describes the behavior that callers can rely on when using `excel-cli capture`. It focuses on the command-line contract and the generated PNG output, not on screenshot implementation details.

## Command Form

`excel-cli capture <file> <sheet> [range]`

- `<file>` is the workbook to open.
- `<sheet>` is the sheet to capture.
- `[range]` optionally restricts the capture area to a specific Excel range.

The command requires both `<file>` and `<sheet>`. If either argument is missing, the command fails with a usage error.

## Input Semantics

### Workbook, Sheet, and Range Selection

`capture` opens the specified workbook, selects the named sheet, and captures either the requested range or the sheet's used range.

If `[range]` is omitted, `capture` determines the used range of the selected sheet and captures that area.

If `[range]` is provided, it is treated as the capture target and is passed through as an Excel-style range string.

### Platform and Backend Expectation

`capture` is intended for environments where the workbook can be opened through the OLE backend. In the current implementation, when the tool falls back to the `excelize` backend, image capture is not supported and the command fails.

For callers, the observable rule is that `capture` requires a runtime environment that supports workbook screenshot capture, which today means the Windows OLE path.

## Successful Behavior

On success, `capture` writes a PNG file under `.excel-cli/` and prints the absolute path of that file to standard output.

Output path pattern:

```text
.excel-cli/capture-<utc-timestamp>-<milliseconds>Z.png
```

If the `.excel-cli` directory does not exist, the command creates it.

Observable characteristics include:

- the output file is a PNG image
- the captured area corresponds to the specified range, or to the used range when no range is supplied
- the workbook itself is not modified as part of capture output generation

## Failure Conditions

`capture` fails when:

- the workbook path cannot be resolved or opened
- the named sheet does not exist
- the used range cannot be determined when `[range]` is omitted
- the selected backend does not support screenshot capture
- the requested range cannot be captured
- the returned image data cannot be decoded
- the PNG output cannot be written
- required arguments are missing

When the command fails, it exits with an error instead of printing an output PNG path.

## Scope Notes

This note defines the caller-visible contract of `capture`: which arguments it accepts, what artifact it produces, and the conditions under which it fails. It does not define how screenshots are obtained internally or behavior for unspecified trailing arguments.

## Related Notes

- [[core-concept]]
- [[list]]
- [[read]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[list]: list "Behavior of the list Command"
[read]: read "Behavior of the read Command"