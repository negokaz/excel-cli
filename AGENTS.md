# Project Guidelines

## Documentation Alignment

- Treat the design notes under `docs/` as the primary source of intent for implementation work in this repository.
- Treat `skills/excel-cli/SKILL.md` as the primary source for agent-facing workflow guidance for this repository, and keep it aligned with the design intent described in `docs/`.
- The files under `docs/` are maintained in Foam format (Markdown with Foam wikilinks and frontmatter), so preserve that structure when editing.
- Start with `docs/index.md` for navigation and `docs/core-concept.md` for cross-cutting behavior and architecture expectations.

## Change Workflow

- For implementation changes, verify whether the affected behavior, output shape, backend choice, or generated artifact is described in `docs/`.
- For changes that affect agent workflow, command usage guidance, or recommended inspection/editing patterns, verify whether `skills/excel-cli/SKILL.md` also needs to be updated.
- If a code change affects documented behavior, update the relevant `docs/` note and `skills/excel-cli/SKILL.md` in the same change when practical. If you cannot update them in the same change, explicitly call out the mismatch and why.
- Do not silently drift away from documented behavior. When docs, code, and tests disagree, surface the discrepancy instead of guessing.
