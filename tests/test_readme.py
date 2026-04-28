from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
README = ROOT / "README.md"


def test_readme_explains_chinese_zotero_layered_workflow():
    text = README.read_text(encoding="utf-8")

    required_phrases = [
        "单 DOCX",
        "中国学位论文",
        "Zotero Group Library",
        "Word 插件",
        "Refresh",
        "GB/T 7714-2015",
        "静态基线",
        "动态引用",
        "uv run python scripts/run_all.py --input examples/demo/raw/thesis.docx",
        "output/doc",
        "output/zotero",
        "output/reports",
    ]
    for phrase in required_phrases:
        assert phrase in text


def test_readme_keeps_public_interface_sanitized_and_single_docx_only():
    text = README.read_text(encoding="utf-8")

    forbidden = [
        "SECOND" + "_DOC",
        "second" + "_doc",
        "chapter2" + "_refs",
        "reference" + "_sources",
        "652" + "8123",
        "/Vol" + "umes/",
        "/Us" + "ers/",
        "面向" + "算力网络",
        "真实" + "主文档",
        "THESIS_REFS_MAIN" + "_DOC",
        "THESIS_REFS_SECOND" + "_DOC",
        "ODF/DOCX " + "Scan",
        "ODF " + "Scan",
        "DOCX " + "Scan",
        "scan" + "_input",
    ]
    for item in forbidden:
        assert item not in text
