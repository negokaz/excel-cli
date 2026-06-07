---
tags:
  - excel-cli
  - concept
  - architecture
---

# Core Concepts of excel-cli

`excel-cli` is a CLI for treating Excel files not as something people open and manipulate in a GUI, but as structured resources that can be addressed and operated on from scripts and other tools.

## In One Sentence

It is a tool for working with Excel workbooks, sheets, and ranges through small, reproducible commands built around paths and verbs.

## Core Ideas

### 1. The Unit of Operation Is file and path

Each `excel-cli` command targets Excel through a concrete pair of `<file> <path>`.
The path identifies a workbook root, sheet, or range precisely, instead of treating a workbook as a vague whole.

This makes it explicit what should happen and where, which simplifies batch processing, machine-readable contracts, and future extension to additional resource kinds.

Initial canonical path examples:

- `/`
- `/Sheet1`
- `/Sheet1/A1`
- `/Sheet1/A1:C3`

### 2. Favor Input and Output Formats That Are Easy to Process

- `read`, `query`, `write`, `add`, and `remove` return JSON on success
- `write` accepts JSON payloads for values, formulas, styles, and properties
- `export` outputs derived artifacts such as HTML and PNG

In particular, range reads and range writes are designed to round-trip cleanly:

- `read --value` returns the same 2-dimensional shape that `write --value` accepts
- `read --formula` returns the same 2-dimensional shape that `write --formula` accepts
- `read --style` returns the same 2-dimensional shape that `write --style` accepts

The point is not to be a CLI mainly for humans to inspect in a terminal, but a CLI whose outputs are easy to pass to the next step in a workflow.

### 3. Choose the Best Backend for Each Runtime Environment

On Windows, if Excel is available, the tool prefers OLE automation so it can use behavior closer to the real Excel application.
When that is not available, it falls back to `excelize`, while keeping the same command structure on macOS and Linux.

With this design, `excel-cli` tries to balance two goals: staying as close as possible to Excel's real behavior and remaining usable across platforms.

### 4. Keep Verbs Small and Predictable

The command surface is organized around a small set of verbs:

- `read`
- `query`
- `write`
- `add`
- `remove`
- `export`

These verbs do not try to embed business logic.
That complexity belongs in the calling script or application, while `excel-cli` acts as a thin boundary around Excel resources and operations.

### 5. Separate Structured Data from Derived Inspection Artifacts

`read` and `query` are for structured data access.
`export` is for derived artifacts such as HTML and PNG that agents and scripts can inspect efficiently.

This separation keeps machine-readable contracts stable while still supporting layout checks, HTML-based exploration, and inspection workflows.

## What This Design Means

The value of this tool is not in replacing Excel entirely, but in turning Excel into something that can be automated as a component.
For that reason, `excel-cli` emphasizes the following qualities:

- Each command has a small, focused responsibility
- Paths identify targets unambiguously
- Outputs are easy to hand off to downstream processing
- Generated artifacts are handled separately from structured JSON
- Backend differences are surfaced only where they are observable and useful, such as the reported backend name or capture support

## Related Notes

- [[index]]

[index]: index "excel-cli Documentation"
