import argparse
from pathlib import Path

from thesis_refs import pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Prepare a Word-plugin operator placeholder DOCX.")
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    parser.add_argument("--source-docx", type=Path, help="Unified DOCX source")
    parser.add_argument("--output-docx", type=Path, help="Operator DOCX output")
    args = parser.parse_args()

    input_docx = pipeline.configure_paths(args.input, args.output_root)
    model = pipeline.build_reference_model(input_docx)
    docx_path, json_path, markdown_path = pipeline.prepare_zotero_operator_docx(
        model,
        source_docx_path=args.source_docx,
        output_docx_path=args.output_docx,
    )
    print(f"docx={docx_path}")
    print(f"json={json_path}")
    print(f"markdown={markdown_path}")


if __name__ == "__main__":
    main()
