---
tags:
  - excel-cli
  - design
  - read
---

# read

This note describes the behavior that callers can rely on when using `excel-cli read`. It focuses on the command-line contract and the generated output, not on backend-specific implementation details.

## Command Form

`excel-cli read <file> <sheet> [--formula] [--style]`

- `<file>` is the workbook to open.
- `<sheet>` is the sheet to read.
- `--formula` shows formulas instead of evaluated values.
- `--style` includes cell style information in the generated HTML.

The command requires both `<file>` and `<sheet>`. If either argument is missing, the command fails with a usage error.

## Input Semantics

### Workbook and Sheet Selection

`read` opens the specified workbook and reads the entire used range of the specified sheet. The caller does not supply a cell range for this command.

The sheet name is treated as a positional argument. Because the command reads `<file>` and `<sheet>` before parsing optional flags, a sheet name that looks like a flag, such as `--style` or `--formula`, is still interpreted as a literal sheet name.

### Output Modes

- Default mode exports displayed cell values.
- `--formula` exports formula text where formulas are present.
- `--style` includes styling information in the HTML output.
- `--formula --style` combines both behaviors.

## Successful Behavior

On success, `read` writes a standalone HTML file under `.excel-cli/` and prints the absolute path of that file to standard output.

Output path pattern:

```
.excel-cli/sheet-<timestamp>.html
```

If the `.excel-cli` directory does not exist, the command creates it.

The generated HTML represents the used range of the sheet as a table. Observable characteristics include:

- a sheet heading
- row and column headers
- cell content rendered into HTML
- line breaks in cell text rendered as `<br>`
- optional CSS when `--style` is enabled

The page also includes visible export metadata, such as the source workbook path, the used range, and the generation time.

## Failure Conditions

`read` fails when:

- the workbook path cannot be resolved or opened
- the named sheet does not exist
- the sheet is treated as empty
- required arguments are missing
- extra arguments are provided
- the HTML output cannot be written

When the command fails, it exits with an error instead of printing an output HTML path.

## Scope Notes

This command is intended for inspection and downstream consumption of sheet content as HTML. This note does not define how the tool computes the used range internally. It defines only the behavior visible to a caller.

## Related Notes

- [[core-concept]]
- [[excel-cli-write-command-behavior]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[excel-cli-write-command-behavior]: excel-cli-write-command-behavior "Behavior of the write Command"