from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
README = ROOT / "README.md"


def test_readme_explains_chinese_zotero_layered_workflow():
    text = README.read_text(encoding="utf-8")

    required_phrases = [
        "单 DOCX",
        "中国学位论文",
        "学位论文",
        "静态整理",
        "可维护引用",
        "后续维护",
        "Zotero Group Library",
        "Word 插件",
        "Refresh",
        "GB/T 7714-2015",
        "uv run python scripts/run_all.py --input examples/demo/raw/thesis.docx",
        "output/doc",
        "output/zotero",
        "output/reports",
    ]
    for phrase in required_phrases:
        assert phrase in text


def test_readme_uses_academic_workflow_tone_instead_of_outsourcing_terms():
    text = README.read_text(encoding="utf-8")

    required_phrases = [
        "使用前知道",
        "脚本会生成新文件",
        "不会就地修改输入的 Word 文档",
        "最后仍需要用 Zotero Word 插件运行 `Refresh`",
    ]
    for phrase in required_phrases:
        assert phrase in text

    forbidden = [
        "交付",
        "全托管",
        "候选",
        "产物目录",
        "审计报告",
        "使用约定",
        "所有生成文件写入",
        "Zotero 字段只能在生成副本中创建",
    ]
    for item in forbidden:
        assert item not in text


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
