import argparse
from pathlib import Path

from thesis_refs import pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the single-DOCX thesis reference pipeline.")
    parser.add_argument("--config", type=Path, help="Optional JSON/TOML config with input_docx and output_root")
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    args = parser.parse_args()

    config = pipeline.load_pipeline_config(args.config)
    input_docx = args.input or config.get("input_docx")
    output_root = args.output_root or config.get("output_root")
    model = pipeline.run_all_exports(input_docx, output_root)
    print(f"references={len(model.references)}")
    print(f"unique_references={len(model.unique_references)}")
    print(f"docx={pipeline.UNIFIED_DOC}")
    print(f"bibtex={pipeline.OUTPUT_DIR / 'zotero' / 'refs.bib'}")
    if model.missing_reference_blocks:
        print(f"missing_reference_blocks={len(model.missing_reference_blocks)}")


if __name__ == "__main__":
    main()
