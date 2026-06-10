---
tags:
  - excel-cli
  - design
  - write
---

# write

This note describes the behavior that callers can rely on when using `excel-cli write`. It defines update channels, input shapes, target restrictions, and failure conditions.

## Command Form

`excel-cli write <file> <path> (--value <json|-> | --formula <json|-> | --style <json|-> | --props <json|->)`

- `<file>` is the workbook to update.
- `<path>` is the canonical target path inside the workbook.
- `--value` writes normal values.
- `--formula` writes formulas for strings that begin with `=` and normal values otherwise.
- `--style` applies cell style updates.
- `--props` updates supported properties of a non-range target.

For each update channel, `-` means the JSON payload is read from standard input.

The command requires `<file>`, `<path>`, and exactly one update channel.

The update channels are mutually exclusive.

## Supported Targets

Initial target support is:

- `--value`: `<sheet>/A1` or `<sheet>/A1:C3`
- `--formula`: `<sheet>/A1` or `<sheet>/A1:C3`
- `--style`: `<sheet>/A1` or `<sheet>/A1:C3`
- `--props`: `<sheet>`

`write` does not support workbook root updates in the initial version.

## Input Contract

### Path

`<path>` must use canonical path syntax.

- the path must not begin with `/`
- sheet names use canonical path segments: Unicode characters are preserved, while ASCII characters that require escaping remain percent-encoded
- range targets are single cells or rectangular ranges only

### `--value`

`--value` accepts a JSON 2-dimensional array, either inline or from standard input when `-` is provided.

- the outer array represents rows
- each inner array represents cells in that row
- the shape must match the target range exactly
- a single cell still requires a 2-dimensional array such as `[[123]]`

### `--formula`

`--formula` accepts a JSON 2-dimensional array, either inline or from standard input when `-` is provided.

- the outer array represents rows
- each inner array represents cells in that row
- the shape must match the target range exactly
- if a string value begins with `=`, it is written as a formula
- otherwise, the item is written as a normal value

This channel is intended to round-trip with `read --formula`.

### `--style`

`--style` accepts a JSON 2-dimensional array, either inline or from standard input when `-` is provided.

- the outer array represents rows
- each inner array represents cells in that row
- the shape must match the target range exactly
- each item must be either a style object or `null`
- `null` means that cell is left unchanged

Each style object may contain:

- `border`
- `font`
- `fill`
- `numFmt`
- `decimalPlaces`

Style enum values and validation rules are the same ones defined by the tool's style schema.

### `--props`

`--props` accepts a JSON object, either inline or from standard input when `-` is provided.

In the initial version, supported target properties are limited to worksheet properties, and the only supported property is:

- `hidden`: boolean

## Successful Behavior

On success, `write` updates the workbook in place, writes formatted JSON to standard output, and exits successfully.

Example:

```json
{
  "path": "Sheet1/A1:C3",
  "kind": "range",
  "action": "write",
  "channel": "value"
}
```

Standard input example:

```sh
echo '[[123]]' | excel-cli write book.xlsx Sheet1/A1 --value -
```

For sheet property updates, `kind` is `sheet` and `channel` is `props`.

## Failure Conditions

`write` fails when:

- the workbook path cannot be resolved or opened
- `<path>` does not follow supported path syntax
- the path kind is not supported for the selected channel
- the target sheet does not exist
- zero or multiple update channels are provided
- `-` is provided but standard input cannot be read
- the selected JSON input cannot be parsed
- a JSON 2-dimensional array is required but not provided
- the input shape does not match the target range
- a style object contains unsupported values
- an unsupported property is specified in `--props`
- required arguments are missing
- the workbook cannot be saved

Callers can expect errors to distinguish at least these cases:

- invalid path syntax
- sheet not found
- unsupported path kind

## Scope Notes

This note defines the external contract of `write`. It does not define backend-specific type coercion or rendering differences beyond the accepted update channels.

## Related Notes

- [[core-concept]]
- [[read]]
- [[add]]
- [[remove]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[read]: read "Behavior of the read Command"
[add]: add "Behavior of the add Command"
[remove]: remove "Behavior of the remove Command"
