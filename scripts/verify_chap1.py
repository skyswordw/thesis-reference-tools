import argparse
from pathlib import Path

from thesis_refs import pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Write a reference verification report for one thesis DOCX.")
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    args = parser.parse_args()

    input_docx = pipeline.configure_paths(args.input, args.output_root)
    model = pipeline.build_reference_model(input_docx)
    pipeline.write_verification_report(model)
    print(f"references_checked={len(model.references)}")
    print(f"report={pipeline.OUTPUT_DIR / 'reports' / 'verify_report.md'}")


if __name__ == "__main__":
    main()
