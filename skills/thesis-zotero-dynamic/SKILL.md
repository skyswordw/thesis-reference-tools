---
name: thesis-zotero-dynamic
description: Use when migrating a thesis Word DOCX from static numeric citations to Zotero dynamic citations, preparing a Zotero Group Library handoff, auditing CSL_CITATION fields, or maintaining GB/T 7714-2015 Word citation refreshability.
---

# Thesis Zotero Dynamic

## Core Rule

Use the validated operator-placeholder route first. Generate Zotero fields only in generated copies, then let Microsoft Word and the Zotero Word plugin prove the document is refreshable before handoff.

## Preconditions

- Do not edit `raw/`; source Word files are read-only.
- Use project-local `uv` commands only.
- Keep all candidate DOCX files under `build/`, `output/`, or `tmp/`.
- Prefer a Zotero Group Library for collaboration, so another computer can refresh citations from the same shared library.

## Default Workflow

1. Build the static baseline:

```bash
uv run python scripts/run_all.py --input examples/demo/raw/thesis.docx
```

2. In Zotero, create or select the Group Library, import `output/zotero/refs.bib`, and confirm citekeys match the reference model, such as `ref001` through the final reference.
3. Before generating or refreshing fields, audit Group Library metadata. Confirm Zotero `itemType` and `language` agree with the project model: report/white paper/research report -> `report`, ITU-T/technical standard -> `standard`, English scholarly items -> `en-US`, Chinese items -> `zh-CN`.
4. Generate the operator copy and checklist:

```bash
uv run python scripts/prepare_zotero_operator_docx.py
```

5. With Zotero running and its local API available, generate the candidate dynamic DOCX:

```bash
uv run python scripts/generate_zotero_field_docx.py
```

6. Install and select the project CSL before final refresh:

```bash
uv run python scripts/install_project_csl.py
```

In Word Zotero Document Preferences, select `NEU Thesis GB/T 7714-2015 Numeric`. This project style follows the school sample more closely than Zotero's default GB/T style: English names are not all-caps, English multi-author entries use `et al.`, Chinese entries may use `等`, journal/conference records use `[J]`/`[C]`, report records use `[R]` or `[R/OL]`, standard records use `[S]` or `[S/OL]`, the reference number is followed by a single normal space, and DOI/URL fields are kept in Zotero metadata but hidden from the Word bibliography.
7. Open the candidate in Word. It must open without repair. Run Zotero Refresh, then use Add/Edit Bibliography to insert or rebuild the bibliography.
8. If the Word bibliography command is unavailable or unreliable, rebuild only the bibliography field result from Zotero Local API plus the selected project CSL. This still uses Zotero item metadata and CSL output, not hand-written semantic text:

```bash
uv run python scripts/rebuild_zotero_bibliography_from_api.py
```

9. Clean the refreshed bibliography residue and paragraph format:

```bash
uv run python scripts/postprocess_zotero_bibliography.py output/doc/论文_zotero_dynamic.docx
```

10. Audit the result against the operator checklist:

```bash
uv run python scripts/audit_zotero_docx.py output/doc/论文_zotero_dynamic.docx --checklist output/zotero/operator_migration_checklist.json --output-dir output/reports/final_dynamic_after_refresh
```

Expected evidence: `CSL_CITATION` fields match citation positions, extracted citation items match checklist occurrences, `csl_bibliography_fields=1`, no static citation markers or operator placeholders remain, no `[[ZOTERO_BIBLIOGRAPHY`/`]]` residue remains, English bibliography entries do not contain Chinese `等`, report/standard entries do not fall back to `[Z]`, and bibliography paragraphs use hanging indent rather than first-line indent.

## Fallback

If generated fields cannot pass Word/Zotero validation, use the operator checklist with the official Zotero Word plugin UI. Replace each `[[ZOTERO_Pxxx]]` with Add/Edit Citation, add all references for a multi-item marker in one citation, and save checkpoint copies every 5-10 positions.

## Safety Checks

- Never treat unverified fields as final.
- Do not write to Zotero SQLite; scripts may copy it for read-only itemID lookup.
- Keep the static baseline deliverable available until the dynamic DOCX passes Word open, Zotero Refresh, bibliography rebuild, and audit.
