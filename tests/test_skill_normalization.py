from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILLS_DIR = ROOT / "skills"
SKILL_NAMES = [
    "thesis-word-format",
    "thesis-zotero-dynamic",
    "thesis-zotero-metadata",
]


def test_skills_have_registry_index():
    index = SKILLS_DIR / "README.md"

    assert index.exists()
    text = index.read_text(encoding="utf-8")
    for name in SKILL_NAMES:
        assert f"skills/{name}/SKILL.md" in text
    assert "single DOCX" in text
    assert "project-local `uv`" in text


def test_project_skills_share_standard_sections_and_boundaries():
    required_sections = [
        "## When To Use",
        "## Inputs And Outputs",
        "## Standard Workflow",
        "## Commands",
        "## Acceptance Checks",
        "## Boundaries",
    ]

    for name in SKILL_NAMES:
        text = (SKILLS_DIR / name / "SKILL.md").read_text(encoding="utf-8")
        for section in required_sections:
            assert section in text, f"{name} missing {section}"
        assert "single DOCX" in text
        assert "project-local `uv`" in text
        assert "Do not edit `raw/`" in text


def test_project_skills_do_not_record_failed_scan_route_or_private_ids():
    combined = "\n".join(
        (SKILLS_DIR / name / "SKILL.md").read_text(encoding="utf-8")
        for name in SKILL_NAMES
    )

    forbidden = [
        "ODF/DOCX " + "Scan",
        "ODF " + "Scan",
        "DOCX " + "Scan",
        "scan" + "_input",
        "652" + "8123",
        "/Vol" + "umes/",
        "/Us" + "ers/",
    ]
    for item in forbidden:
        assert item not in combined
