from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL_DIR = ROOT / "skills" / "thesis-zotero-dynamic"


def test_zotero_dynamic_skill_bundle_exists_and_covers_verified_route():
    skill_path = SKILL_DIR / "SKILL.md"
    openai_yaml_path = SKILL_DIR / "agents" / "openai.yaml"

    assert skill_path.exists()
    assert openai_yaml_path.exists()

    skill_text = skill_path.read_text(encoding="utf-8")
    openai_yaml = openai_yaml_path.read_text(encoding="utf-8")

    for required in [
        "Group Library",
        "refs.bib",
        "itemType",
        "language",
        "report",
        "standard",
        "prepare_zotero_operator_docx.py",
        "generate_zotero_field_docx.py",
        "audit_zotero_docx.py",
        "Word",
        "Zotero Refresh",
        "Add/Edit Bibliography",
        "GB/T 7714-2015",
        "NEU Thesis GB/T 7714-2015 Numeric",
        "install_project_csl.py",
        "postprocess_zotero_bibliography.py",
        "CSL_CITATION",
        "csl_bibliography_fields",
    ]:
        assert required in skill_text

    assert "Thesis Zotero Dynamic" in openai_yaml
    assert "Word dynamic Zotero citations" in openai_yaml


def test_zotero_dynamic_skill_excludes_failed_scan_route():
    skill_text = (SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")

    for forbidden in [
        "ODF/DOCX " + "Scan",
        "ODF " + "Scan",
        "Pan" + "doc",
        "scan" + "_input",
    ]:
        assert forbidden not in skill_text


def test_zotero_dynamic_skill_records_safety_boundaries():
    skill_text = (SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")
    agents_text = (ROOT / "AGENTS.md").read_text(encoding="utf-8")

    for required in [
        "Do not edit `raw/`",
        "project-local `uv`",
        "generated copies",
        "open without repair",
        "unverified fields",
        "Zotero Word plugin",
    ]:
        assert required in skill_text

    assert "skills/thesis-zotero-dynamic/SKILL.md" in agents_text
    assert "Zotero dynamic citations" in agents_text
