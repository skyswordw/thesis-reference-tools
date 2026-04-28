from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import urllib.parse
import urllib.request
import zipfile
from html import unescape
from html.parser import HTMLParser
from pathlib import Path

from lxml import etree

from thesis_refs import pipeline
from thesis_refs.pipeline import bibtex_key
from thesis_refs.zotero_metadata import ZOTERO_GROUP_ID_ENV, require_group_id


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": WORD_NS}
ANCHOR = "[[ZOTERO_BIBLIOGRAPHY]]"
STYLE_ID = "neu-thesis-gbt7714-2015-numeric"
LOCAL_API_BASE = "http://127.0.0.1:23119"


def default_group_id() -> int | None:
    value = os.environ.get(ZOTERO_GROUP_ID_ENV)
    return int(value) if value else None


class CslEntryParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.entries: list[str] = []
        self._depth = 0
        self._chunks: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        classes = " ".join(value or "" for key, value in attrs if key == "class")
        if tag == "div" and "csl-entry" in classes.split():
            self._depth = 1
            self._chunks = []
        elif self._depth:
            self._depth += 1

    def handle_endtag(self, tag: str) -> None:
        if not self._depth:
            return
        self._depth -= 1
        if self._depth == 0:
            self.entries.append(normalize_space("".join(self._chunks)))
            self._chunks = []

    def handle_data(self, data: str) -> None:
        if self._depth:
            self._chunks.append(data)


def normalize_space(text: str) -> str:
    return " ".join(unescape(text or "").split())


def api_json(url: str) -> object:
    with urllib.request.urlopen(url, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def api_text(url: str) -> str:
    with urllib.request.urlopen(url, timeout=30) as response:
        return response.read().decode("utf-8")


def citation_key_from_extra(extra: object) -> str | None:
    if not isinstance(extra, str):
        return None
    for line in extra.splitlines():
        if line.lower().startswith(("citation key:", "citation-key:")):
            return line.split(":", 1)[1].strip()
    return None


def fetch_group_item_keys(group_id: int, api_base: str) -> dict[str, str]:
    item_keys: dict[str, str] = {}
    start = 0
    while True:
        query = urllib.parse.urlencode(
            {"format": "json", "include": "data", "limit": 100, "start": start}
        )
        batch = api_json(f"{api_base.rstrip('/')}/api/groups/{group_id}/items?{query}")
        if not isinstance(batch, list) or not batch:
            break
        for item in batch:
            if not isinstance(item, dict):
                continue
            data = item.get("data")
            if not isinstance(data, dict):
                continue
            citation_key = data.get("citationKey") or citation_key_from_extra(data.get("extra"))
            item_key = data.get("key")
            if citation_key and item_key:
                item_keys[str(citation_key)] = str(item_key)
        start += len(batch)
    return item_keys


def fetch_bibliography_entry(group_id: int, item_key: str, api_base: str, style: str) -> str:
    query = urllib.parse.urlencode({"format": "bib", "style": style})
    html = api_text(f"{api_base.rstrip('/')}/api/groups/{group_id}/items/{item_key}?{query}")
    parser = CslEntryParser()
    parser.feed(html)
    if not parser.entries:
        raise RuntimeError(f"No CSL bibliography entry returned for Zotero item {item_key}")
    return parser.entries[0]


def bibliography_entries(group_id: int, api_base: str, style: str, input_docx: Path | None = None, output_root: Path | None = None) -> list[str]:
    resolved_input = pipeline.configure_paths(input_docx, output_root)
    model = pipeline.build_reference_model(resolved_input)
    item_keys = fetch_group_item_keys(group_id, api_base)
    entries: list[str] = []
    for ref in model.unique_references:
        citekey = bibtex_key(ref)
        item_key = item_keys.get(citekey)
        if not item_key:
            raise RuntimeError(f"Missing Zotero item for citekey {citekey}")
        entry = fetch_bibliography_entry(group_id, item_key, api_base, style)
        entry = re.sub(r"^\[\d+\]", f"[{ref.global_id}]", entry)
        entries.append(entry)
    return entries


def w_el(local_name: str, text: str | None = None, **attrs: str) -> etree._Element:
    element = etree.Element(f"{{{WORD_NS}}}{local_name}")
    for key, value in attrs.items():
        if key == "xml_space":
            element.set(f"{{{XML_NS}}}space", value)
        else:
            element.set(f"{{{WORD_NS}}}{key}", value)
    if text is not None:
        element.text = text
    return element


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))


def bibliography_ppr() -> etree._Element:
    ppr = w_el("pPr")
    ppr.append(w_el("pStyle", val="Bibliography"))
    tabs = w_el("tabs")
    tabs.append(w_el("tab", val="left", pos="480"))
    ppr.append(tabs)
    ppr.append(w_el("spacing", before="0", after="0", line="300", lineRule="auto"))
    ppr.append(w_el("ind", left="480", hanging="480"))
    ppr.append(w_el("jc", val="left"))
    rpr = w_el("rPr")
    rpr.append(w_el("rFonts", ascii="Times New Roman", hAnsi="Times New Roman", eastAsia="SimSun", cs="Times New Roman"))
    rpr.append(w_el("sz", val="24"))
    rpr.append(w_el("szCs", val="24"))
    ppr.append(rpr)
    return ppr


def text_run(text: str) -> etree._Element:
    run = w_el("r")
    rpr = w_el("rPr")
    rpr.append(w_el("rFonts", ascii="Times New Roman", hAnsi="Times New Roman", eastAsia="SimSun", cs="Times New Roman"))
    rpr.append(w_el("sz", val="24"))
    rpr.append(w_el("szCs", val="24"))
    run.append(rpr)
    text_element = w_el("t", text)
    if text[:1].isspace() or text[-1:].isspace():
        text_element.set(f"{{{XML_NS}}}space", "preserve")
    run.append(text_element)
    return run


def field_run(kind: str | None = None, instruction: str | None = None) -> etree._Element:
    run = text_run("")
    for child in list(run):
        if child.tag == f"{{{WORD_NS}}}t":
            run.remove(child)
    if kind:
        run.append(w_el("fldChar", fldCharType=kind))
    if instruction is not None:
        run.append(w_el("instrText", instruction, xml_space="preserve"))
    return run


def bibliography_paragraphs(entries: list[str]) -> list[etree._Element]:
    paragraphs: list[etree._Element] = []
    instruction = ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '
    for index, entry in enumerate(entries):
        paragraph = w_el("p")
        paragraph.append(bibliography_ppr())
        if index == 0:
            paragraph.append(field_run("begin"))
            paragraph.append(field_run(instruction=instruction))
            paragraph.append(field_run("separate"))
        paragraph.append(text_run(entry))
        if index == len(entries) - 1:
            paragraph.append(field_run("end"))
        paragraphs.append(paragraph)
    return paragraphs


def replace_bibliography_anchor(input_docx: Path, output_docx: Path, entries: list[str]) -> Path:
    with zipfile.ZipFile(input_docx, "r") as zin:
        document_xml = zin.read("word/document.xml")
        root = etree.fromstring(document_xml)
        anchor_paragraph = None
        for paragraph in root.xpath("//w:p", namespaces=NS):
            if paragraph_text(paragraph).strip() == ANCHOR:
                anchor_paragraph = paragraph
                break
        if anchor_paragraph is None:
            raise RuntimeError(f"Bibliography anchor not found: {ANCHOR}")
        parent = anchor_paragraph.getparent()
        if parent is None:
            raise RuntimeError("Bibliography anchor has no parent")
        insert_at = parent.index(anchor_paragraph)
        parent.remove(anchor_paragraph)
        for offset, paragraph in enumerate(bibliography_paragraphs(entries)):
            parent.insert(insert_at + offset, paragraph)
        updated_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        temp_path = output_docx.with_name(f".{output_docx.stem}.tmp.docx")
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/document.xml":
                    data = updated_xml
                zout.writestr(info, data)
    shutil.move(temp_path, output_docx)
    return output_docx


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Rebuild a Zotero bibliography field result from Zotero Local API + project CSL."
    )
    parser.add_argument("--input-docx", type=Path, default=Path("output/doc/论文_zotero_dynamic_generated_candidate.docx"))
    parser.add_argument("--output-docx", type=Path, default=Path("output/doc/论文_zotero_dynamic.docx"))
    parser.add_argument("--group-id", type=int, default=default_group_id())
    parser.add_argument("--api-base", default=LOCAL_API_BASE)
    parser.add_argument("--style", default=STYLE_ID)
    parser.add_argument("--input", type=Path, help="Complete thesis DOCX to process")
    parser.add_argument("--output-root", type=Path, help="Output directory root")
    args = parser.parse_args()

    group_id = require_group_id(args.group_id)
    entries = bibliography_entries(group_id, args.api_base, args.style, args.input, args.output_root)
    output = replace_bibliography_anchor(args.input_docx, args.output_docx, entries)
    print(f"bibliography_entries={len(entries)}")
    print(f"output_docx={output}")


if __name__ == "__main__":
    main()
