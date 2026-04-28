from __future__ import annotations

import argparse
import json
from pathlib import Path

import httpx

from thesis_refs.zotero_metadata import (
    DEFAULT_GROUP_ID,
    DEFAULT_LOCAL_API_BASE,
    DEFAULT_WEB_API_BASE,
    MetadataRecord,
    api_key_from_env,
    apply_patch_to_zotero_web,
    audit_zotero_metadata_items,
    creators_from_responsibility,
    enrich_by_doi,
    fetch_zotero_group_items,
    make_patch_entry,
    project_item_type,
    project_language,
    require_group_id,
    write_metadata_patch_report,
)
from thesis_refs import pipeline
from thesis_refs.pipeline import bibtex_key, reference_container_fields


def fallback_record_from_reference(ref) -> MetadataRecord:
    container = reference_container_fields(ref)
    item_type = project_item_type(ref.ref_type, ref.title or ref.raw_text)
    record = MetadataRecord(
        title=ref.title or ref.raw_text,
        creators=creators_from_responsibility(ref.responsibility, include_others=False),
        year=ref.year,
        date=ref.year,
        doi=ref.doi,
        url=ref.url,
        item_type=item_type,
        container_title=container.get("booktitle") or container.get("journal"),
        volume=container.get("volume"),
        issue=container.get("number"),
        pages=container.get("pages"),
        publisher=container.get("publisher"),
        language=project_language(ref.title or ref.raw_text, ref.responsibility, item_type),
        source="project_reference",
    )
    return record


def merge_records(primary: MetadataRecord | None, fallback: MetadataRecord | None) -> MetadataRecord | None:
    if primary is None:
        return fallback
    if fallback is None:
        return primary
    return MetadataRecord(
        title=primary.title or fallback.title,
        creators=primary.creators or fallback.creators,
        year=primary.year or fallback.year,
        date=primary.date or fallback.date,
        doi=primary.doi or fallback.doi,
        url=primary.url or fallback.url,
        # The project reference model is authoritative for Zotero itemType and
        # language because public DOI APIs can classify reports/standards as
        # generic articles.
        item_type=fallback.item_type or primary.item_type,
        container_title=primary.container_title or fallback.container_title,
        volume=primary.volume or fallback.volume,
        issue=primary.issue or fallback.issue,
        pages=primary.pages or fallback.pages,
        publisher=primary.publisher or fallback.publisher,
        language=fallback.language or primary.language,
        source=primary.source or fallback.source,
    )


def build_patch(group_id: int, api_base: str, input_docx: Path | None = None, output_root: Path | None = None) -> list[dict[str, object]]:
    items = fetch_zotero_group_items(group_id, api_base)
    audit = audit_zotero_metadata_items(items)
    suspicious_keys = {
        row.get("item_key")
        for row in audit["items"]
        if row.get("item_key") and row.get("doi")
    }
    resolved_input = pipeline.configure_paths(input_docx, output_root)
    model = pipeline.build_reference_model(resolved_input)
    fallback_by_key = {bibtex_key(ref): fallback_record_from_reference(ref) for ref in model.unique_references}
    patch: list[dict[str, object]] = []
    with httpx.Client(timeout=30) as client:
        for item in items:
            data = item.get("data", {})
            if not isinstance(data, dict):
                continue
            doi = data.get("DOI")
            citation_key = data.get("citationKey")
            if not citation_key:
                for line in str(data.get("extra") or "").splitlines():
                    if line.lower().startswith(("citation key:", "citation-key:")):
                        citation_key = line.split(":", 1)[1].strip()
                        break
            fallback = fallback_by_key.get(str(citation_key))
            if data.get("key") not in suspicious_keys and not fallback:
                continue
            record = enrich_by_doi(str(doi), client) if doi else None
            record = merge_records(record, fallback)
            if not record:
                continue
            entry = make_patch_entry(item, record)
            if entry:
                patch.append(entry)
    return patch


def main() -> None:
    parser = argparse.ArgumentParser(description="Enrich Zotero Group Library metadata from public DOI metadata APIs.")
    parser.add_argument("--group-id", type=int, default=DEFAULT_GROUP_ID)
    parser.add_argument("--local-api-base", default=DEFAULT_LOCAL_API_BASE)
    parser.add_argument("--web-api-base", default=DEFAULT_WEB_API_BASE)
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    parser.add_argument("--env-file", type=Path, default=Path(".env.local"))
    parser.add_argument("--output-dir", type=Path, default=Path("output/reports"))
    parser.add_argument("--apply", action="store_true", help="Apply the generated patch through Zotero Web API.")
    args = parser.parse_args()

    group_id = require_group_id(args.group_id)
    patch = build_patch(group_id, args.local_api_base, args.input, args.output_root)
    json_path, md_path = write_metadata_patch_report(patch, args.output_dir)
    print(f"patch_items={len(patch)}")
    print(f"json={json_path}")
    print(f"markdown={md_path}")

    if args.apply:
        api_key = api_key_from_env(args.env_file)
        if not api_key:
            raise SystemExit(
                "ZOTERO_API_KEY not found. Put it in project-local .env.local before using --apply."
            )
        results = apply_patch_to_zotero_web(patch, group_id, api_key, args.web_api_base)
        apply_log = args.output_dir / "zotero_metadata_apply_log.json"
        apply_log.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"applied_items={len(results)}")
        print(f"apply_log={apply_log}")


if __name__ == "__main__":
    main()
