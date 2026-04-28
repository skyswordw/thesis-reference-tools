---
name: thesis-word-format
description: Use when formatting Chinese doctoral thesis Word/DOCX deliverables, applying Northeastern University thesis layout guidance, GB/T 7714 sequential references, final reference lists, inline citation superscripts, or Chinese narrative citations such as 文献[n].
---

# Thesis Word Format

## Core Workflow

1. Treat the Word file as the 博士学位论文 deliverable. Preserve DOCX structure and verify the rendered pages before calling it final.
2. Read `references/degree-format-summary.md` before changing page layout, headings, body fonts, figure/table captions, or reference-list typography.
3. Read `references/docx-format-execution-checklist.md` before making whole-document Word formatting changes.
4. Read `references/citation-position-rules.md` before editing inline citations.
5. Use `scripts/normalize_citation_markers.py` after reference renumbering to set terminal citation markers as Word superscript runs (上标).
6. If `soffice` and Poppler are available, render DOCX -> PDF -> PNG and inspect pages. If they are missing, report that visual layout still needs local Word/LibreOffice review.

## Citation Rules

- Use sequential numeric references: final reference entries are `[n] Author. Title[type]. ...`.
- For this project, dynamic Zotero bibliographies should use `NEU Thesis GB/T 7714-2015 Numeric`, not the stock Zotero GB/T style, so English authors remain title case, DOI/URL metadata stays in Zotero for verification, and the Word bibliography hides DOI/URL fields to match the school examples.
- Bibliography 语义 format belongs to Zotero metadata plus CSL, not final-text patching. Fix `itemType`, `language`, author metadata, and the project CSL when `[R]`/`[S]`, `et al.`, `等`, DOI, or URL behavior is wrong.
- After Zotero Refresh, run `scripts/postprocess_zotero_bibliography.py` only for layout and residue cleanup: remove bibliography anchor residue, restore hanging indent/line spacing, and keep Word field structure intact.
- For citations used as evidence at the end of a clause or sentence, keep `[n]` attached to the sentence and set the marker as superscript.
- For citations used as the grammatical subject at the start of a Chinese sentence, write `文献[n]指出/提出/认为...` or `作者在文献[n]中...`; keep `文献`/`作者在文献` baseline, but set the `[n]` marker as superscript.
- Avoid English-paper phrasing like `Sun等[n]...` or `Chen等[n]...` as a Chinese sentence opener. Rewrite to `文献[n]...` unless the author name is semantically necessary.
- When multiple references support one statement, use one marker such as `[1,3-5]`, not repeated adjacent markers.

## Tooling

```bash
uv run python skills/thesis-word-format/scripts/normalize_citation_markers.py input.docx output.docx --report output/reports/citation_marker_report.json
```

The script styles all in-text numeric markers as superscript, rewrites English author-led sentence/clause openings such as `Chen等[n]...` to `文献[n]...`, and reports the rewrite count. It does not replace careful reading; inspect the report before final export.

Read the risk boundary in `references/docx-format-execution-checklist.md` before running the script on documents with Zotero fields, hyperlinks, footnotes, text boxes, bookmarks, or comments.
