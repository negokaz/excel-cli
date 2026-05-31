# Project Guidelines

## Documentation Alignment

- Treat the design notes under `docs/` as the primary source of intent for implementation work in this repository.
- Start with `docs/index.md` for navigation and `docs/core-concept.md` for cross-cutting behavior and architecture expectations.

## Change Workflow

- For implementation changes, verify whether the affected behavior, output shape, backend choice, or generated artifact is described in `docs/`.
- If a code change affects documented behavior, update the relevant document in the same change when practical. If you cannot update docs in the same change, explicitly call out the mismatch and why.
- Do not silently drift away from documented behavior. When docs, code, and tests disagree, surface the discrepancy instead of guessing.
