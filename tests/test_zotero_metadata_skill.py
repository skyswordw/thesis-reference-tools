from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL_PATH = ROOT / "skills" / "thesis-zotero-metadata" / "SKILL.md"


def test_zotero_metadata_skill_file_exists():
    assert SKILL_PATH.exists()


def test_zotero_metadata_skill_covers_group_library_web_api_word_refresh_and_audit():
    skill_text = SKILL_PATH.read_text(encoding="utf-8")

    for required in [
        "Group Library",
        "Web API",
        "Word Refresh",
        "audit",
        "itemType",
        "report",
        "standard",
        "language",
        "et al.",
        "non-DOI",
    ]:
        assert required in skill_text


def test_zotero_metadata_skill_excludes_failed_scan_routes():
    skill_text = SKILL_PATH.read_text(encoding="utf-8")

    for forbidden in [
        "ODF/DOCX " + "Scan",
        "ODF " + "Scan",
        "DOCX " + "Scan",
    ]:
        assert forbidden not in skill_text
