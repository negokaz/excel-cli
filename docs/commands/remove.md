---
tags:
  - excel-cli
  - design
  - remove
---

# remove

This note describes the behavior that callers can rely on when using `excel-cli remove`. It defines the deletion contract, dry-run behavior, and failure conditions.

## Command Form

`excel-cli remove <file> <path> [--force]`

- `<file>` is the workbook to update.
- `<path>` is the canonical path to remove.
- `--force` performs the deletion. Without it, the command runs as a dry-run validation.

The command requires `<file>` and `<path>`.

## Supported Targets

In the initial version, `remove` supports only worksheet deletion.

The supported form is:

- `/<sheet>`

## Successful Behavior

### Dry-Run Mode

When `--force` is not provided, `remove` does not modify the workbook. It only checks whether the target could be removed and reports that result as JSON.

Example:

```json
{
  "path": "/Sales",
  "kind": "sheet",
  "action": "remove",
  "wouldRemove": true
}
```

### Forced Removal

When `--force` is provided, `remove` deletes the target sheet, saves the workbook in place, writes formatted JSON to standard output, and exits successfully.

Example:

```json
{
  "path": "/Sales",
  "kind": "sheet",
  "action": "remove"
}
```

## Failure Conditions

`remove` fails when:

- the workbook path cannot be resolved or opened
- `<path>` does not follow supported path syntax
- `<path>` is not a supported removable path in the initial version
- the target is the workbook's only worksheet
- the target sheet does not exist
- required arguments are missing
- the workbook cannot be saved in forced mode

Callers can expect errors to distinguish at least these cases:

- invalid path syntax
- sheet not found
- unsupported path kind

## Scope Notes

This note defines the initial external contract of `remove`. It does not define recursive deletion, workbook deletion, or removal of tables, charts, or named ranges.

## Related Notes

- [[core-concept]]
- [[add]]
- [[write]]

[core-concept]: core-concept "Core Concepts of excel-cli"
[add]: add "Behavior of the add Command"
[write]: write "Behavior of the write Command"
