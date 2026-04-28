import importlib
import importlib.util
from pathlib import Path
from types import SimpleNamespace

import pytest


def zotero_metadata():
    try:
        return importlib.import_module("thesis_refs.zotero_metadata")
    except ModuleNotFoundError as exc:
        raise AssertionError(
            "Expected thesis_refs.zotero_metadata to provide parse_display_authors, "
            "map_crossref_work, and bibtex_author_value"
        ) from exc


def test_parse_display_author_completes_initials_to_structured_creators():
    metadata = zotero_metadata()

    creators = metadata.parse_display_authors(
        "Yang D, Gong K, Ren J, et al",
        known_creators=[
            {"given": "Dong", "family": "Yang"},
            {"given": "Kai", "family": "Gong"},
            {"given": "Jian", "family": "Ren"},
            {"given": "Xiaoping", "family": "Chen"},
        ],
    )

    assert creators[:3] == [
        {"creatorType": "author", "firstName": "Dong", "lastName": "Yang"},
        {"creatorType": "author", "firstName": "Kai", "lastName": "Gong"},
        {"creatorType": "author", "firstName": "Jian", "lastName": "Ren"},
    ]
    assert creators[3] == {"creatorType": "author", "fieldMode": 1, "lastName": "others"}


def test_crossref_fixture_maps_ieee_network_article_metadata():
    metadata = zotero_metadata()

    crossref_work = {
        "DOI": "10.1109/mnet.007.2100444",
        "container-title": ["IEEE Network"],
        "volume": "36",
        "issue": "2",
        "page": "16-24",
        "published-print": {"date-parts": [[2022, 3]]},
        "type": "journal-article",
        "title": [
            "Intelligent Task Offloading for 6G Low-Latency Edge Intelligence"
        ],
        "author": [
            {"given": "Dong", "family": "Yang"},
            {"given": "Kai", "family": "Gong"},
            {"given": "Jian", "family": "Ren"},
        ],
    }

    item = metadata.map_crossref_work(crossref_work)

    assert item["DOI"] == "10.1109/mnet.007.2100444"
    assert item["publicationTitle"] == "IEEE Network"
    assert item["volume"] == "36"
    assert item["issue"] == "2"
    assert item["pages"] == "16-24"
    assert item["itemType"] == "journalArticle"


def test_bibtex_author_value_uses_zotero_friendly_last_first_names():
    metadata = zotero_metadata()

    creators = [
        {"creatorType": "author", "firstName": "Dong", "lastName": "Yang"},
        {"creatorType": "author", "firstName": "Kai", "lastName": "Gong"},
        {"creatorType": "author", "firstName": "Jian", "lastName": "Ren"},
        {"creatorType": "author", "fieldMode": 1, "lastName": "others"},
    ]

    assert metadata.bibtex_author_value(creators) == (
        "Yang, Dong and Gong, Kai and Ren, Jian and others"
    )
    assert "Yang D and Gong K" not in metadata.bibtex_author_value(creators)


def test_project_reference_heuristics_map_report_standard_and_papers():
    metadata = zotero_metadata()

    assert metadata.project_item_type("R/OL", "中国综合算力评价白皮书（2023年）") == "report"
    assert metadata.project_item_type("R", "区域算力网: 高速互联篇研究报告") == "report"
    assert metadata.project_item_type(
        "EB/OL",
        "中国移动发布《云智算技术白皮书》: 定义体系架构，提出十大方向",
    ) == "report"
    assert metadata.project_item_type("EB/OL", "Computing Power Network Research Report") == "report"
    assert metadata.project_item_type(
        "S/OL",
        "Recommendation ITU-T Q.4144 (04/2025): Signalling requirements",
    ) == "standard"
    assert metadata.project_item_type(
        "EB/OL",
        "ITU-T Y.2500-series - Computing power networks standardization roadmap",
    ) == "standard"
    assert metadata.project_item_type("J", "Journal paper") == "journalArticle"
    assert metadata.project_item_type("C", "Conference paper") == "conferencePaper"


def test_make_patch_entry_updates_non_doi_report_with_language():
    metadata = zotero_metadata()

    item = {
        "version": 7,
        "data": {
            "key": "ABC123",
            "itemType": "document",
            "title": "中国综合算力评价白皮书（2023年）",
            "creators": [],
            "extra": "Citation Key: ref008",
        },
    }
    record = metadata.MetadataRecord(
        title="中国综合算力评价白皮书（2023年）",
        item_type="report",
        language="zh-CN",
        source="project_reference",
    )

    entry = metadata.make_patch_entry(item, record)

    assert entry is not None
    assert entry["doi"] is None
    assert entry["changed_fields"] == ["itemType", "language"]
    assert entry["updated_data"]["itemType"] == "report"
    assert entry["updated_data"]["language"] == "zh-CN"


def test_make_patch_entry_sets_english_paper_language():
    metadata = zotero_metadata()

    item = {
        "data": {
            "key": "DEF456",
            "itemType": "journalArticle",
            "title": "Intelligent Task Offloading for 6G Low-Latency Edge Intelligence",
            "extra": "Citation Key: ref026",
        }
    }
    record = metadata.MetadataRecord(
        title="Intelligent Task Offloading for 6G Low-Latency Edge Intelligence",
        item_type="journalArticle",
        language="en-US",
        source="project_reference",
    )

    entry = metadata.make_patch_entry(item, record)

    assert entry is not None
    assert entry["changed_fields"] == ["language"]
    assert entry["updated_data"]["language"] == "en-US"


def test_audit_reports_expected_item_type_mismatch():
    metadata = zotero_metadata()

    audit = metadata.audit_zotero_metadata_items(
        [
            {
                "data": {
                    "key": "ABC123",
                    "itemType": "document",
                    "title": "Recommendation ITU-T Q.4144 (04/2025): Signalling requirements",
                    "extra": "Citation Key: ref007",
                }
            }
        ],
        expected_by_citekey={"ref007": {"item_type": "standard"}},
    )

    assert audit["summary"]["expected_item_type_mismatches"] == 1
    assert audit["items"][0]["expected_item_type"] == "standard"
    assert audit["items"][0]["item_type"] == "document"


def test_enrich_fallback_record_maps_non_doi_project_report():
    script_path = Path(__file__).resolve().parents[1] / "scripts" / "enrich_zotero_metadata.py"
    spec = importlib.util.spec_from_file_location("enrich_zotero_metadata", script_path)
    assert spec and spec.loader
    enrich = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(enrich)

    record = enrich.fallback_record_from_reference(
        SimpleNamespace(
            title="区域算力网: 高速互联篇研究报告",
            raw_text="国家信息中心, 华为. 区域算力网: 高速互联篇研究报告[R/OL]. 2025-08-29[2026-04-18].",
            responsibility="国家信息中心, 华为",
            ref_type="R/OL",
            year="2025",
            doi=None,
            url="https://example.test/report.pdf",
        )
    )

    assert record.item_type == "report"
    assert record.language == "zh-CN"
    assert record.doi is None


def test_enrich_merge_keeps_project_item_type_and_language_authoritative():
    script_path = Path(__file__).resolve().parents[1] / "scripts" / "enrich_zotero_metadata.py"
    spec = importlib.util.spec_from_file_location("enrich_zotero_metadata", script_path)
    assert spec and spec.loader
    enrich = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(enrich)
    metadata = zotero_metadata()

    primary = metadata.MetadataRecord(
        title="Computing Power Network Research Report",
        item_type="journalArticle",
        language="en-US",
        container_title="External Metadata Source",
        source="crossref",
    )
    fallback = metadata.MetadataRecord(
        title="Computing Power Network Research Report",
        item_type="report",
        language="zh-CN",
        source="project_reference",
    )

    record = enrich.merge_records(primary, fallback)

    assert record is not None
    assert record.item_type == "report"
    assert record.language == "zh-CN"
    assert record.container_title == "External Metadata Source"
