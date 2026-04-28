import argparse
from pathlib import Path

from thesis_refs import pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Deduplicate and renumber references from one thesis DOCX.")
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    args = parser.parse_args()

    input_docx = pipeline.configure_paths(args.input, args.output_root)
    model = pipeline.build_reference_model(input_docx)
    pipeline.write_ref_map(model)
    duplicate_count = len(model.references) - len(model.unique_references)
    print(f"references={len(model.references)}")
    print(f"unique_references={len(model.unique_references)}")
    print(f"merged_duplicates={duplicate_count}")


if __name__ == "__main__":
    main()
