import argparse
from pathlib import Path

from thesis_refs import pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Export GB/T 7714 reference text and Zotero import files.")
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    args = parser.parse_args()

    input_docx = pipeline.configure_paths(args.input, args.output_root)
    model = pipeline.build_reference_model(input_docx)
    pipeline.write_final_references(model)
    pipeline.write_bibtex(model)
    pipeline.write_csl_json(model)
    pipeline.write_zotero_readme(model)
    print(f"unique_references={len(model.unique_references)}")
    print(f"refs_final={pipeline.BUILD_DIR / 'refs_final.txt'}")
    print(f"bibtex={pipeline.OUTPUT_DIR / 'zotero' / 'refs.bib'}")


if __name__ == "__main__":
    main()
