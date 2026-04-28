# Thesis Reference Tools

Tools for extracting, verifying, renumbering, and exporting references from a single thesis DOCX.

## Official Interface

The supported input is one thesis `.docx` file. The pipeline discovers citation markers and reference material from that document and writes generated artifacts under `output/`.

Demo usage:

```bash
uv run python scripts/run_all.py --input examples/demo/raw/thesis.docx
```

You can also set `THESIS_REFS_INPUT_DOCX=/path/to/thesis.docx` or pass a JSON/TOML config containing only `input_docx`, `output_root`, and `zotero_group_id`.

This publishable version documents the single-DOCX workflow only.

## Zotero Dynamic Migration Boundary

The static pipeline creates plain-text citation numbers and Zotero import files. Final handoff documents should prefer the official Word + Zotero plugin workflow. Generated dynamic-citation fields are allowed only in experiment/checkpoint copies under `build/`, `output/`, or `tmp/`, and must be validated by opening in Microsoft Word without repair prompts, running Zotero Refresh, rebuilding the bibliography, and checking the citation/reference mapping before handoff.

Useful generated outputs and follow-up tools:

- `output/zotero/README_zotero.md` for the handoff workflow.
- `output/zotero/migration_checklist.md` for paragraph-by-paragraph citation replacement.
- `scripts/prepare_zotero_operator_docx.py` to create a unique-placeholder operator copy for official UI fallback.
- `output/zotero/group_library_handoff.md` for the recommended Zotero Group Library setup.
- `scripts/audit_zotero_docx.py` to check whether a DOCX contains real Zotero fields after pilot or batch migration.
- `styles/neu-thesis-gbt7714-2015-numeric.csl` for the project-local Zotero style that matches the school sample more closely than the stock GB/T style.
- `scripts/postprocess_zotero_bibliography.py` to clean Zotero bibliography residue, normalize reference-list hanging indent/line spacing, and optionally restore pre-refresh media compression after Word Refresh.

## Thesis Word Formatting

For Word thesis formatting and citation-marker cleanup, use the project-local skill in `skills/thesis-word-format/`.
