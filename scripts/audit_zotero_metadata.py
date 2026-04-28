from __future__ import annotations

import argparse
import json
from pathlib import Path

from thesis_refs.zotero_metadata import (
    DEFAULT_GROUP_ID,
    DEFAULT_LOCAL_API_BASE,
    audit_zotero_metadata_items,
    fetch_zotero_group_items,
    project_item_type,
    require_group_id,
)
from thesis_refs import pipeline
from thesis_refs.pipeline import bibtex_key


def write_report(audit: dict[str, object], output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    json_path = output_dir / "zotero_metadata_audit.json"
    md_path = output_dir / "zotero_metadata_audit.md"
    json_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")

    summary = audit["summary"]
    lines = [
        "# Zotero 元数据审计",
        "",
        f"- Items scanned: {summary['items_scanned']}",
        f"- Citekeys: {summary['citekeys']}",
        f"- Duplicate citekeys: {summary['duplicate_citekeys']}",
        f"- Suspicious items: {summary['suspicious_items']}",
        f"- Inverted initial creator items: {summary['inverted_initial_creator_items']}",
        f"- Fake others creator items: {summary['fake_others_creator_items']}",
        f"- DOI items missing container: {summary['doi_items_missing_container']}",
        f"- Expected itemType mismatches: {summary['expected_item_type_mismatches']}",
        "",
        "| Citekey | Item | Actual type | Expected type | DOI | Inverted creators | Fake others | Missing container | Type mismatch | Title |",
        "|---|---|---|---|---|---:|---:|---|---|---|",
    ]
    for row in audit["items"]:
        title = str(row.get("title") or "").replace("|", "/")
        lines.append(
            f"| {row.get('citation_key') or ''} | {row.get('item_key') or ''} | "
            f"{row.get('item_type') or ''} | {row.get('expected_item_type') or ''} | "
            f"{row.get('doi') or ''} | {row.get('inverted_initial_creators') or 0} | "
            f"{row.get('fake_others_creators') or 0} | {row.get('missing_container')} | "
            f"{row.get('item_type_mismatch')} | {title} |"
        )
    md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return json_path, md_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit Zotero Group Library metadata for thesis references.")
    parser.add_argument("--group-id", type=int, default=DEFAULT_GROUP_ID)
    parser.add_argument("--api-base", default=DEFAULT_LOCAL_API_BASE)
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    parser.add_argument("--output-dir", type=Path, default=Path("output/reports"))
    args = parser.parse_args()

    group_id = require_group_id(args.group_id)
    input_docx = pipeline.configure_paths(args.input, args.output_root)
    items = fetch_zotero_group_items(group_id, args.api_base)
    model = pipeline.build_reference_model(input_docx)
    expected_by_citekey = {
        bibtex_key(ref): project_item_type(ref.ref_type, ref.title or ref.raw_text)
        for ref in model.unique_references
    }
    audit = audit_zotero_metadata_items(items, expected_by_citekey=expected_by_citekey)
    json_path, md_path = write_report(audit, args.output_dir)
    summary = audit["summary"]
    print(f"items_scanned={summary['items_scanned']}")
    print(f"suspicious_items={summary['suspicious_items']}")
    print(f"inverted_initial_creator_items={summary['inverted_initial_creator_items']}")
    print(f"fake_others_creator_items={summary['fake_others_creator_items']}")
    print(f"doi_items_missing_container={summary['doi_items_missing_container']}")
    print(f"expected_item_type_mismatches={summary['expected_item_type_mismatches']}")
    print(f"json={json_path}")
    print(f"markdown={md_path}")


if __name__ == "__main__":
    main()
