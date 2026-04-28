import argparse
from pathlib import Path

from thesis_refs import pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a unified-reference DOCX from one complete thesis DOCX.")
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    parser.add_argument("--output-docx", type=Path, help="Optional output DOCX path")
    args = parser.parse_args()

    input_docx = pipeline.configure_paths(args.input, args.output_root)
    model = pipeline.build_reference_model(input_docx)
    output_path = pipeline.rewrite_main_docx(model, args.output_docx)
    print(f"docx={output_path}")
    print(f"unique_references={len(model.unique_references)}")


if __name__ == "__main__":
    main()
