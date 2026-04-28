from __future__ import annotations

import argparse
import importlib.util
import json
import shutil
import subprocess
import sys
from pathlib import Path

from thesis_refs import pipeline
from thesis_refs.zotero_metadata import DEFAULT_GROUP_ID, require_group_id


def load_normalize_docx():
    script = Path("skills/thesis-word-format/scripts/normalize_citation_markers.py")
    spec = importlib.util.spec_from_file_location("normalize_citation_markers", script)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load {script}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.normalize_docx


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Rebuild the formatted operator DOCX after Zotero metadata updates."
    )
    parser.add_argument("--static-docx", type=Path, default=Path("output/doc/论文_统一编号.docx"))
    parser.add_argument("--formatted-docx", type=Path, default=Path("output/doc/论文_统一编号_引用上标.docx"))
    parser.add_argument("--operator-docx", type=Path, default=Path("output/doc/论文_zotero_dynamic_operator.docx"))
    parser.add_argument("--candidate-docx", type=Path, default=Path("output/doc/论文_zotero_dynamic_generated_candidate.docx"))
    parser.add_argument("--checklist", type=Path, default=Path("output/zotero/operator_migration_checklist.json"))
    parser.add_argument("--report", type=Path, default=Path("output/reports/citation_marker_report.json"))
    parser.add_argument("--group-id", type=int, default=DEFAULT_GROUP_ID)
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    parser.add_argument("--skip-field-generation", action="store_true")
    args = parser.parse_args()

    input_docx = pipeline.configure_paths(args.input, args.output_root)
    model = pipeline.build_reference_model(input_docx)

    pipeline.rewrite_main_docx(model, args.static_docx, input_docx)
    report = load_normalize_docx()(args.static_docx, args.formatted_docx)
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    pipeline.prepare_zotero_operator_docx(
        model,
        source_docx_path=args.formatted_docx,
        output_docx_path=args.operator_docx,
        output_dir=Path("output/zotero"),
    )
    if args.operator_docx.exists():
        checkpoint = Path("output/doc/论文_zotero_dynamic_metadata_ready.docx")
        shutil.copyfile(args.operator_docx, checkpoint)
        print(f"checkpoint={checkpoint}")
    if not args.skip_field_generation:
        group_id = require_group_id(args.group_id)
        subprocess.run(
            [
                sys.executable,
                "scripts/generate_zotero_field_docx.py",
                "--source-docx",
                str(args.operator_docx),
                "--checklist",
                str(args.checklist),
                "--output-docx",
                str(args.candidate_docx),
                "--group-id",
                str(group_id),
            ],
            check=True,
        )
        print(f"candidate_docx={args.candidate_docx}")
    print(f"static_docx={args.static_docx}")
    print(f"formatted_docx={args.formatted_docx}")
    print(f"operator_docx={args.operator_docx}")
    print("next=open candidate in Word, run Zotero Refresh, rebuild bibliography, and save output/doc/论文_zotero_dynamic.docx")


if __name__ == "__main__":
    main()
