from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CSL_PATH = ROOT / "styles" / "neu-thesis-gbt7714-2015-numeric.csl"


def test_neu_thesis_csl_exists_and_has_project_identity():
    assert CSL_PATH.exists()
    text = CSL_PATH.read_text(encoding="utf-8")

    assert "NEU Thesis GB/T 7714-2015 Numeric" in text
    assert "neu-thesis-gbt7714-2015-numeric" in text


def test_neu_thesis_csl_keeps_school_author_case_and_hides_access_fields():
    text = CSL_PATH.read_text(encoding="utf-8")

    assert 'name-part name="family" text-case="uppercase"' not in text
    assert '<if type="article-journal paper-conference" match="none">' in text
    assert '<text variable="DOI" prefix="DOI:"/>' not in text
    assert '<text variable="URL"/>' not in text
    assert "Keep DOI/URL in Zotero metadata" in text
