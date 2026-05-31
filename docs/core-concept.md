---
tags:
  - excel-cli
  - concept
  - architecture
---

# Core Concepts of excel-cli

`excel-cli` is a CLI for treating Excel files not as something people open and manipulate in a GUI, but as stable targets that can be called from scripts and other tools.

## In One Sentence

It is a tool for working with Excel workbooks, sheets, and ranges as small, reproducible commands.

## Core Ideas

### 1. The Unit of Operation Is file, sheet, and range

Each `excel-cli` command targets Excel through concrete units such as `<file> <sheet> <range>` instead of treating a workbook as a vague whole.
This makes it explicit what should happen and where, which simplifies batch processing and integration with other tools.

### 2. Favor Input and Output Formats That Are Easy to Process

- `list` returns JSON
- `write` and `format` accept two-dimensional JSON arrays
- `read` outputs HTML so you can inspect not only values but also layout and structure
- `capture` outputs PNG

In particular, `write` and `format` are designed so that the specified range matches the shape of the JSON array, which keeps the update target unambiguous.

The point is not to be a CLI mainly for humans to inspect in a terminal, but a CLI whose outputs are easy to pass to the next step in a workflow.

### 3. Choose the Best Backend for Each Runtime Environment

On Windows, if Excel is available, the tool prefers OLE automation so it can use behavior closer to the real Excel application.
When that is not available, it falls back to `excelize`, while keeping the same command structure on macOS and Linux.

With this design, `excel-cli` tries to balance two goals: staying as close as possible to Excel's real behavior and remaining usable across platforms.

### 4. Keep Subcommands Primitive

The available commands are limited to the basic operations `new`, `list`, `read`, `write`, `format`, and `capture`; they do not try to embed complex business logic.
That complexity belongs in the calling script or application, while `excel-cli` acts as a thin boundary around Excel operations.

## What This Design Means

The value of this tool is not in replacing Excel entirely, but in turning Excel into something that can be automated as a component.
For that reason, `excel-cli` emphasizes the following qualities:

- Each command has a small, focused responsibility
- Arguments have clear meaning
- Outputs are easy to hand off to downstream processing
- Backend differences are encapsulated internally

## Related Notes

- [[index]]

[index]: index "excel-cli Documentation"
