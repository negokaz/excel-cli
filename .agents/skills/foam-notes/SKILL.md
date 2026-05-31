---
name: foam-notes
description: Work with Foam note repositories through Foam CLI. Covers linting, listing, note inspection and lifecycle, outlines, links, tags, grep/search, and rename workflows.
---

# Foam Notes

Work with Foam note-taking workspaces in VS Code through Foam CLI. [Foam](https://foamnotes.com) is a free, open-source personal knowledge management system built on standard Markdown files with wikilinks.

Use `--workspace <dir>` to target another workspace, or set `FOAM_WORKSPACE` before running commands.

## Quick Reference

- **Wikilinks**: `[[note-name]]` connect notes bidirectionally.
- **Embeds**: `![[note-name]]` include note content inline.
- **Backlinks**: `foam links NOTE --incoming` shows references to a note.
- **Tags**: Use inline `#tag` or frontmatter `tags: [tag1, tag2]`.
- **Workspace override**: `foam ... --workspace path/to/notes`

## Core Concepts

### Wikilinks

Create connections between notes using double brackets:

```markdown
See also [[related-concept]] for more information.
```

- Autocomplete with `[[` and start typing the note name.
- Navigate with `Ctrl+Click` (or `Cmd+Click` on Mac).
- Create new notes by following non-existent links.
- Link to sections with `[[note-name#Section Title]]`.

### Backlinks

Backlinks show which notes reference the current note.

From the CLI:

```bash
foam links current-note --incoming
```

In VS Code:
- Use the Connections view to inspect related notes.
- Switch between incoming, outgoing, or combined link views.

### Tags

Organize notes beyond wikilinks:

```markdown
# Inline tags
#machine-learning #research #in-progress

# Frontmatter tags
---
tags: [machine-learning, research, in-progress]
---
```

- Hierarchical tags such as `#programming/python` are supported.
- Use `foam tag list` to inventory tags.
- Use `foam tag search TAG` or `foam search --tag TAG` to find tagged notes.

## Foam CLI Workflows

All commands support `--workspace <dir>` to override the workspace root.

### lint

Check the workspace for structural issues and optionally auto-fix what Foam can repair:

```bash
foam lint
foam lint --fix
foam lint --rule missing-heading
foam lint --format json
```

Use this before bulk refactors or in CI. Exit code `0` means no issues, `2` means issues were found, and `1` means the command failed.

### list

List notes, tags, and common graph hygiene buckets:

```bash
foam list notes
foam list notes --tag project --tag active
foam list notes --type note
foam list tags --sort count
foam list orphans
foam list deadends
foam list placeholders
```

Use `foam list tags` for tag inventory and `foam list placeholders` to find unresolved wikilinks.

### note

Inspect a note, resolve its identifier, create notes, or remove notes from the workspace:

```bash
foam note show project-plan
foam note show project-plan --format json
foam note id project-plan
foam note create --title "Project Plan"
foam note delete project-plan
```

Use `foam note move NOTE` when you want a file move or rename that also rewrites wikilinks. Use `--path` when exact file targeting is more reliable than identifier resolution.

### outline

Print the heading structure of a note when you need a quick table of contents:

```bash
foam outline project-plan
foam outline project-plan --format json
```

This is the CLI equivalent of checking a note's section structure before linking or renaming headings.

### links

Inspect incoming and outgoing links for a note:

```bash
foam links project-plan
foam links project-plan --incoming
foam links project-plan --outgoing
```

Use `--incoming` for backlinks and `--outgoing` to audit what a note references. `foam connections` is an alias for the same command family.

### tag

Work with tags directly from the CLI:

```bash
foam tag list
foam tag list --format json
foam tag search project
foam tag rename old-tag new-tag
foam tag rename old-tag new-tag --force
```

`foam tag search TAG` is the direct replacement for tag lookup scripts, and `foam tag rename` updates hierarchical children as well.

### grep

Search raw note content with a regular expression or plain text pattern:

```bash
foam grep "TODO"
foam grep "TODO" --context 2
foam grep "TODO" --limit 50
foam grep "TODO" --no-line-number
```

Use `grep` when the match can appear anywhere in note content and does not need graph-aware lookup.

### search

Search notes by title, alias, tag, or frontmatter property:

```bash
foam search "meeting"
foam search --tag project
foam search --tag project --tag active
foam search --property status=draft
foam search --type note
```

Use `search` when you want Foam-aware filtering instead of raw text matching. `--tag` is repeatable and uses AND semantics.

### rename

Refactor notes and link targets while preserving workspace integrity:

```bash
foam rename note old-note new-note
foam rename note old-note new-note --target-path archive/
foam rename tag old-tag new-tag
foam rename section project-plan "Old Heading" "New Heading"
foam rename block project-plan old-anchor new-anchor
```

Use this for any rename that should also rewrite wikilinks across the workspace. `foam rename note` is the safest replacement for manual file renames.

## Common Tasks

### Create Or Inspect A Note

```bash
foam note create --title "Research Topic"
foam note show research-topic
foam note id research-topic
```

### Find Relationships

```bash
foam links research-topic --incoming
foam links research-topic --outgoing
foam outline research-topic
```

### Search By Tag Or Text

```bash
foam tag search research
foam search --tag research --property status=active
foam grep "next step"
```

### Rename Safely

```bash
foam rename note old-topic new-topic
foam rename section old-topic "Old Heading" "New Heading"
foam rename tag old-tag new-tag
```

## External Resources

- **Official site**: https://foamnotes.com
- **GitHub**: https://github.com/foambubble/foam
- **Discord**: https://foambubble.github.io/join-discord/w

## Tips

1. **Start small**: Foam works best with consistent note-taking habits.
2. **Link liberally**: Create wikilinks even to notes you have not written yet.
3. **Prefer CLI-safe refactors**: Use `foam rename ...` and `foam note move ...` instead of manual file renames when links matter.
4. **Choose the right search**: Use `foam search` for note metadata and `foam grep` for raw content.
5. **Keep it standard**: Foam uses standard Markdown, so your notes remain portable.