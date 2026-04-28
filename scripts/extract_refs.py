import argparse
from pathlib import Path

from thesis_refs import pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract local reference blocks from one complete thesis DOCX.")
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    args = parser.parse_args()

    input_docx = pipeline.configure_paths(args.input, args.output_root)
    model = pipeline.build_reference_model(input_docx)
    pipeline.write_raw_refs(model)
    print(f"reference_blocks={len({ref.source_section for ref in model.references})}")
    print(f"references={len(model.references)}")
    print(f"citation_occurrences={len(model.citation_occurrences)}")
    print(f"missing_reference_blocks={len(model.missing_reference_blocks)}")


if __name__ == "__main__":
    main()
