from __future__ import annotations

import argparse
from pathlib import Path

from thesis_refs.pipeline import audit_zotero_docx, optimize_docx_media, postprocess_zotero_bibliography_docx


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Clean Zotero bibliography residue and normalize thesis reference-list paragraph format."
    )
    parser.add_argument("input_docx", type=Path)
    parser.add_argument("output_docx", type=Path, nargs="?")
    parser.add_argument("--checklist", type=Path, default=Path("output/zotero/operator_migration_checklist.json"))
    parser.add_argument("--report-dir", type=Path, default=Path("output/reports/final_dynamic_custom_csl"))
    parser.add_argument(
        "--media-source",
        type=Path,
        help="Optional DOCX whose word/media files should replace Word-rewritten media in the output.",
    )
    args = parser.parse_args()

    output = postprocess_zotero_bibliography_docx(args.input_docx, args.output_docx)
    if args.media_source:
        output = optimize_docx_media(output, args.media_source, output)
    audit = audit_zotero_docx(output, args.report_dir, args.checklist)
    print(f"output_docx={output}")
    for key in (
        "complex_field_unclosed_count",
        "bibliography_field_unclosed_count",
        "visible_field_code_text_count",
        "bibliography_anchor_residue_count",
        "stray_closing_bracket_paragraphs",
        "bibliography_reference_paragraphs",
        "bibliography_first_line_indent_count",
        "bibliography_hanging_indent_count",
        "bibliography_tight_line_spacing_count",
        "bibliography_non_left_alignment_count",
        "bibliography_access_field_samples",
    ):
        value = audit[key]
        if isinstance(value, list):
            value = len(value)
        print(f"{key}={value}")


if __name__ == "__main__":
    main()
