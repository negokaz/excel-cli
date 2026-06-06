---
tags:
  - excel-cli
  - design
  - format
---

# format

This note describes the behavior that callers can rely on when using `excel-cli format`. It focuses on the command contract, accepted style schema, and failure conditions.

## Command Form

`excel-cli format <file> <sheet> <range> <styles>`

- `<file>` is the workbook to update.
- `<sheet>` is the target sheet name.
- `<range>` is the target Excel range.
- `<styles>` is a JSON 2-dimensional array of cell style objects or `null`.

The command requires four positional arguments. It updates an existing workbook in place.

## Input Contract

### Range

The range is expressed separately from the sheet name. Inputs such as `Sheet1!A1:C3` are rejected. Valid inputs are standard Excel-style cell references and rectangular ranges such as `A1`, `A1:C2`, or `$A$1:$C$2`.

### Styles

`<styles>` must be a JSON 2-dimensional array.

- the outer array represents rows
- each inner array represents cells in that row
- the shape must match the target range exactly
- each item must be either a style object or `null`
- `null` means that cell is left unchanged

### Style Object

Each style object may contain these top-level properties:

- `border`: array of border definitions
- `font`: font settings
- `fill`: fill settings
- `numFmt`: custom number format string
- `decimalPlaces`: integer from `0` to `30`

Supported border fields:

- `type`: one of `left`, `right`, `top`, `bottom`, `diagonalDown`, `diagonalUp`
- `style`: one of `none`, `continuous`, `dash`, `dot`, `double`, `dashDot`, `dashDotDot`, `slantDashDot`, `mediumDashDot`, `mediumDashDotDot`
- `color`: `#RRGGBB`

Supported font fields:

- `bold`: boolean
- `italic`: boolean
- `underline`: one of `none`, `single`, `double`, `singleAccounting`, `doubleAccounting`
- `size`: integer from `1` to `409`
- `strike`: boolean
- `color`: `#RRGGBB`
- `vertAlign`: one of `baseline`, `superscript`, `subscript`

Supported fill fields:

- `type`: `gradient` or `pattern`
- `pattern`: one of `none`, `solid`, `mediumGray`, `darkGray`, `lightGray`, `darkHorizontal`, `darkVertical`, `darkDown`, `darkUp`, `darkGrid`, `darkTrellis`, `lightHorizontal`, `lightVertical`, `lightDown`, `lightUp`, `lightGrid`, `lightTrellis`, `gray125`, `gray0625`
- `color`: array of `#RRGGBB`
- `shading`: one of `horizontal`, `vertical`, `diagonalDown`, `diagonalUp`, `fromCenter`, `fromCorner`

Example:

```json
[
  [
    {
      "font": { "bold": true, "color": "#FF0000" },
      "fill": { "type": "pattern", "pattern": "solid", "color": ["#FFF2CC"] }
    },
    null
  ]
]
```

## Successful Behavior

On success, `format` applies the specified styles to the target range, saves the workbook in place, and exits successfully.

Observable behavior includes:

- only cells represented by non-`null` items are updated
- cells mapped to `null` keep their existing style
- the workbook file is modified in place

Unlike `read` and `capture`, this command does not create a generated artifact under `.excel-cli/`.

## Failure Conditions

`format` fails when:

- the workbook path cannot be resolved or opened
- the target sheet does not exist
- the range is invalid or includes a sheet name
- `<styles>` is not valid JSON
- `<styles>` is not a JSON 2-dimensional array
- the row count does not match the range height
- any row's column count does not match the range width
- a style object contains an unsupported enum value
- a color is not in `#RRGGBB` format
- `font.size` is outside `1` to `409`
- `decimalPlaces` is outside `0` to `30`
- required arguments are missing
- extra arguments are provided

When the command fails, it exits with an error. Success is signaled by the updated workbook file and a successful process exit status.

## Scope Notes

This note defines the external contract of `format`: how callers specify a target range, how style JSON is shaped, and which inputs are rejected. It does not define backend-specific rendering differences inside Excel itself.

## Related Notes

- [[core-concept]]
- [[write]]
- [[read]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[write]: write "Behavior of the write Command"
[read]: read "Behavior of the read Command"
