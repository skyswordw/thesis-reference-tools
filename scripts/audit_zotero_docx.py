import argparse
from pathlib import Path

from thesis_refs.pipeline import audit_zotero_docx


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit Zotero fields in a DOCX file.")
    parser.add_argument("docx", type=Path, help="DOCX file to audit")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output/reports"),
        help="Directory for the audit report",
    )
    parser.add_argument(
        "--checklist",
        type=Path,
        default=None,
        help="Optional Zotero migration checklist for sequence validation",
    )
    args = parser.parse_args()

    audit = audit_zotero_docx(args.docx, args.output_dir, args.checklist)

    for key in (
        "zotero_item_fields",
        "csl_citation_fields",
        "csl_bibliography_fields",
        "extracted_citation_items",
        "pandoc_citation_markers",
        "pandoc_citation_items",
        "operator_placeholders",
        "remaining_static_citation_markers",
        "bibliography_anchor_residue_count",
        "stray_closing_bracket_paragraphs",
        "bibliography_reference_paragraphs",
        "bibliography_first_line_indent_count",
        "bibliography_hanging_indent_count",
        "bibliography_tight_line_spacing_count",
        "bibliography_non_left_alignment_count",
        "bibliography_access_field_samples",
        "bibliography_english_chinese_et_al_count",
        "bibliography_report_standard_z_type_count",
        "bibliography_report_standard_type_mismatch_count",
    ):
        if key in audit:
            value = audit[key]
            if isinstance(value, list):
                value = len(value)
            print(f"{key}={value}")


if __name__ == "__main__":
    main()
