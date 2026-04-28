# Project Skills

These project-level skills document reusable workflows for a single DOCX thesis reference pipeline. Use project-local `uv` commands, keep `raw/` read-only, and write generated artifacts under `build/`, `output/`, or `tmp/`.

| Skill | Use When | Entry |
|---|---|---|
| Thesis Word Format | Word thesis layout, citation-marker 上标, `文献[n]` cleanup, and reference-list layout checks | `skills/thesis-word-format/SKILL.md` |
| Thesis Zotero Dynamic | Static numeric citations need Zotero dynamic Word fields, Group Library handoff, Word Refresh, or field audit | `skills/thesis-zotero-dynamic/SKILL.md` |
| Thesis Zotero Metadata | Zotero records need author, DOI/container, `itemType`, language, report/standard, or `et al.` cleanup before Refresh | `skills/thesis-zotero-metadata/SKILL.md` |

## Routing

- Start with `skills/thesis-word-format/SKILL.md` for visible DOCX formatting and citation marker styling.
- Start with `skills/thesis-zotero-metadata/SKILL.md` when bibliography semantics are wrong because Zotero metadata is incomplete or misclassified.
- Start with `skills/thesis-zotero-dynamic/SKILL.md` when replacing static citations with Zotero Word fields or auditing dynamic field integrity.

## Baseline Command

```bash
uv run python scripts/run_all.py --input examples/demo/raw/thesis.docx
```
