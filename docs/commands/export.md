---
tags:
  - excel-cli
  - design
  - export
---

# export

This note describes the behavior that callers can rely on when using `excel-cli export`. It defines export formats, target support, generated artifacts, and failure conditions.

## Command Form

`excel-cli export <file> <path> --format <html|png> [options]`

- `<file>` is the workbook to open.
- `<path>` is the canonical target path inside the workbook.
- `--format` selects the exported artifact type and is required.

The command requires `<file>`, `<path>`, and `--format`.

## Supported Targets

In the initial version, `export` supports:

- `<sheet>`
- `<sheet>/A1`
- `<sheet>/A1:C3`

Workbook root export is not supported in the initial version.

## Output Behavior

On success, `export` writes a derived artifact under `.excel-cli/`, prints the absolute path of that artifact to standard output, and exits successfully.

The output location is always managed by the tool. Callers do not provide an output path.

Observable characteristics include:

- HTML exports produce `.html` files
- PNG exports produce `.png` files
- the workbook itself is not modified as part of export generation

The intent of `export` is to produce alternate representations that agents and scripts can inspect efficiently, such as HTML for structured exploration or PNG for visual layout checks.

## HTML Export

Form:

`excel-cli export <file> <path> --format html [--formula] [--style]`

Options:

- `--formula` exports formulas instead of displayed values where applicable
- `--style` includes style information in the generated HTML

These options affect HTML export only.

For a range target, the exported area is the specified range. For a sheet target, the exported area is the sheet's used range.

## PNG Export

Form:

`excel-cli export <file> <path> --format png`

PNG export has no additional initial-version options. It captures the current visible appearance of the selected sheet or range.

As an observable contract, PNG export requires a runtime environment and backend combination that supports screenshot capture.

## Failure Conditions

`export` fails when:

- the workbook path cannot be resolved or opened
- `<path>` does not follow supported path syntax
- the path kind is not supported for export
- the target sheet does not exist
- `--format` is missing or unsupported
- HTML output cannot be generated or written
- PNG capture is not supported by the selected backend or runtime
- PNG data cannot be produced or written
- required arguments are missing

Callers can expect errors to distinguish at least these cases:

- invalid path syntax
- sheet not found
- unsupported path kind

## Scope Notes

This note defines the external contract of `export`. It does not define internal HTML structure, screenshot implementation details, or backend-specific visual differences beyond the generated artifact type and visible target area.

## Related Notes

- [[core-concept]]
- [[read]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[read]: read "Behavior of the read Command"
