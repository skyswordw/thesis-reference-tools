import json
import argparse
from pathlib import Path

from thesis_refs import pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a static-to-Zotero migration checklist.")
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    parser.add_argument("--target-docx", type=Path, help="Unified DOCX to scan")
    args = parser.parse_args()

    input_docx = pipeline.configure_paths(args.input, args.output_root)
    model = pipeline.build_reference_model(input_docx)
    json_path, markdown_path = pipeline.write_zotero_migration_checklist(model, target_docx_path=args.target_docx)

    print(f"json={json_path}")
    print(f"markdown={markdown_path}")

    try:
        summary = json.loads(json_path.read_text(encoding="utf-8")).get("summary", {})
    except OSError:
        summary = {}

    for key in (
        "unique_references",
        "citation_positions",
        "citation_paragraphs",
        "covered_citation_occurrences",
        "expected_citation_occurrences",
        "paragraph_coverage_ok",
        "paragraph_coverage_mismatches",
    ):
        if key in summary:
            print(f"{key}={summary[key]}")


if __name__ == "__main__":
    main()
