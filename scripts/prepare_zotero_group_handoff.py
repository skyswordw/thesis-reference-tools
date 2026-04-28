import argparse
from pathlib import Path

from thesis_refs import pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Write Zotero Group Library handoff notes.")
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    args = parser.parse_args()

    input_docx = pipeline.configure_paths(args.input, args.output_root)
    model = pipeline.build_reference_model(input_docx)
    output_path = pipeline.write_zotero_group_handoff(model)
    print(f"handoff={output_path}")


if __name__ == "__main__":
    main()
