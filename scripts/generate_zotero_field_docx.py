import argparse
import json
import os
import random
import re
import shutil
import sqlite3
import string
import urllib.request
import zipfile
from pathlib import Path

from lxml import etree


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": WORD_NS}
PLACEHOLDER_RE = re.compile(r"\[\[ZOTERO_P\d{3}\]\]")
ZOTERO_GROUP_ID_ENV = "ZOTERO_GROUP_ID"


def default_group_id() -> int | None:
    value = os.environ.get(ZOTERO_GROUP_ID_ENV)
    return int(value) if value else None


def require_group_id(group_id: int | None) -> int:
    if group_id is None:
        raise SystemExit(f"Provide --group-id or set {ZOTERO_GROUP_ID_ENV}.")
    return int(group_id)


def api_json(url: str):
    with urllib.request.urlopen(url, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def fetch_group_items(group_id: int, api_base: str) -> dict[str, dict[str, object]]:
    items: dict[str, dict[str, object]] = {}
    start = 0
    while True:
        batch = api_json(
            f"{api_base}/api/groups/{group_id}/items?format=json&include=data&limit=100&start={start}"
        )
        if not batch:
            break
        for item in batch:
            data = item.get("data", {})
            citation_key = data.get("citationKey") or _citation_key_from_extra(data.get("extra", ""))
            item_key = data.get("key")
            if citation_key and item_key:
                items[str(citation_key)] = {"item_key": str(item_key), "api_data": data}
        start += len(batch)
    return items


def _citation_key_from_extra(extra: object) -> str | None:
    if not isinstance(extra, str):
        return None
    for line in extra.splitlines():
        if line.lower().startswith("citation key:"):
            return line.split(":", 1)[1].strip()
        if line.lower().startswith("citation-key:"):
            return line.split(":", 1)[1].strip()
    return None


def fetch_csl_item(group_id: int, item_key: str, api_base: str) -> dict[str, object]:
    data = api_json(f"{api_base}/api/groups/{group_id}/items/{item_key}?format=csljson")
    if not data:
        raise RuntimeError(f"No CSL JSON returned for Zotero item {item_key}")
    item = data[0]
    if not isinstance(item, dict):
        raise RuntimeError(f"Unexpected CSL JSON for Zotero item {item_key}")
    return item


def copy_zotero_sqlite(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)


def load_item_ids(sqlite_path: Path, item_keys: set[str]) -> dict[str, int]:
    if not item_keys:
        return {}
    placeholders = ",".join("?" for _ in item_keys)
    with sqlite3.connect(sqlite_path) as conn:
        rows = conn.execute(
            f"select key, itemID from items where key in ({placeholders})",
            sorted(item_keys),
        ).fetchall()
    return {str(key): int(item_id) for key, item_id in rows}


def marker_plain_text(marker: str) -> str:
    return marker.replace(" ", "")


def citation_id() -> str:
    alphabet = string.ascii_letters + string.digits
    return "".join(random.choice(alphabet) for _ in range(8))


def build_payload(
    group_id: int,
    marker: str,
    zotero_keys: list[str],
    ref_to_item: dict[str, dict[str, object]],
    item_ids: dict[str, int],
    csl_cache: dict[str, dict[str, object]],
) -> dict[str, object]:
    citation_items = []
    for ref_key in zotero_keys:
        item_meta = ref_to_item[ref_key]
        item_key = str(item_meta["item_key"])
        item_id = item_ids[item_key]
        csl_item = dict(csl_cache[ref_key])
        csl_item["id"] = item_id
        csl_item.setdefault("citation-key", ref_key)
        citation_items.append(
            {
                "id": item_id,
                "uris": [f"http://zotero.org/groups/{group_id}/items/{item_key}"],
                "itemData": csl_item,
            }
        )
    plain = marker_plain_text(marker)
    return {
        "citationID": citation_id(),
        "properties": {
            "unsorted": False,
            "formattedCitation": f"\\super {plain}\\nosupersub{{}}",
            "plainCitation": plain,
            "noteIndex": 0,
        },
        "citationItems": citation_items,
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
    }


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


def clone_run_properties(template_run: etree._Element) -> etree._Element | None:
    props = template_run.find("w:rPr", namespaces=NS)
    return etree.fromstring(etree.tostring(props)) if props is not None else None


def field_runs(template_run: etree._Element, instruction: str, display_text: str) -> list[etree._Element]:
    props = clone_run_properties(template_run)

    begin = w_el("r")
    begin.append(w_el("fldChar", fldCharType="begin"))

    instr = w_el("r")
    if props is not None:
        instr.append(etree.fromstring(etree.tostring(props)))
    instr.append(w_el("instrText", instruction, xml_space="preserve"))

    separate = w_el("r")
    separate.append(w_el("fldChar", fldCharType="separate"))

    result = w_el("r")
    result_props = clone_run_properties(template_run)
    if result_props is None:
        result_props = w_el("rPr")
    result_props.append(w_el("vertAlign", val="superscript"))
    result.append(result_props)
    result.append(w_el("t", display_text))

    end = w_el("r")
    end.append(w_el("fldChar", fldCharType="end"))

    return [begin, instr, separate, result, end]


def text_run(template_run: etree._Element, text: str) -> etree._Element:
    run = w_el("r")
    props = clone_run_properties(template_run)
    if props is not None:
        run.append(props)
    text_element = w_el("t", text)
    if text[:1].isspace() or text[-1:].isspace():
        text_element.set(f"{{{XML_NS}}}space", "preserve")
    run.append(text_element)
    return run


def replace_placeholders(
    root: etree._Element,
    positions: list[dict[str, object]],
    group_id: int,
    ref_to_item: dict[str, dict[str, object]],
    item_ids: dict[str, int],
    csl_cache: dict[str, dict[str, object]],
) -> int:
    by_placeholder = {f"[[ZOTERO_{pos['position_id']}]]": pos for pos in positions}
    replaced = 0
    for text_node in root.xpath("//w:t[contains(text(), 'ZOTERO_P')]", namespaces=NS):
        text = text_node.text or ""
        matches = list(PLACEHOLDER_RE.finditer(text))
        if not matches:
            continue
        template_run = text_node.getparent()
        parent = template_run.getparent()
        index = parent.index(template_run)
        new_runs: list[etree._Element] = []
        cursor = 0
        for match in matches:
            if match.start() > cursor:
                new_runs.append(text_run(template_run, text[cursor : match.start()]))
            placeholder = match.group(0)
            pos = by_placeholder.get(placeholder)
            if not pos:
                new_runs.append(text_run(template_run, placeholder))
                cursor = match.end()
                continue
            marker = str(pos["marker_text"])
            zotero_keys = [str(key) for key in pos["zotero_keys"]]
            payload = build_payload(group_id, marker, zotero_keys, ref_to_item, item_ids, csl_cache)
            instruction = " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(
                payload, ensure_ascii=False, separators=(",", ":")
            ) + " "
            new_runs.extend(field_runs(template_run, instruction, marker_plain_text(marker)))
            replaced += 1
            cursor = match.end()
        if cursor < len(text):
            new_runs.append(text_run(template_run, text[cursor:]))
        parent.remove(template_run)
        for offset, run in enumerate(new_runs):
            parent.insert(index + offset, run)
    return replaced


def write_docx_with_document_xml(source_docx: Path, output_docx: Path, document_xml: bytes) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source_docx, "r") as zin, zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                data = document_xml
            zout.writestr(info, data)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate an experimental Zotero-field DOCX from operator placeholders.")
    parser.add_argument("--source-docx", type=Path, default=Path("output/doc/论文_zotero_dynamic_operator.docx"))
    parser.add_argument("--checklist", type=Path, default=Path("output/zotero/operator_migration_checklist.json"))
    parser.add_argument("--output-docx", type=Path, default=Path("output/doc/论文_zotero_dynamic_generated_candidate.docx"))
    parser.add_argument("--group-id", type=int, default=default_group_id())
    parser.add_argument("--api-base", default="http://127.0.0.1:23119")
    parser.add_argument("--zotero-sqlite", type=Path, default=Path.home() / "Zotero" / "zotero.sqlite")
    parser.add_argument("--sqlite-copy", type=Path, default=Path("build/zotero.sqlite.copy"))
    args = parser.parse_args()
    group_id = require_group_id(args.group_id)

    checklist = json.loads(args.checklist.read_text(encoding="utf-8"))
    positions = checklist["positions"]
    expected_refs = {str(key) for pos in positions for key in pos["zotero_keys"]}

    ref_to_item = fetch_group_items(group_id, args.api_base)
    missing_refs = sorted(expected_refs - set(ref_to_item))
    if missing_refs:
        raise SystemExit(f"Missing Zotero items for citekeys: {', '.join(missing_refs)}")

    item_keys = {str(ref_to_item[ref]["item_key"]) for ref in expected_refs}
    copy_zotero_sqlite(args.zotero_sqlite, args.sqlite_copy)
    item_ids = load_item_ids(args.sqlite_copy, item_keys)
    missing_ids = sorted(item_keys - set(item_ids))
    if missing_ids:
        raise SystemExit(f"Missing Zotero numeric itemIDs for item keys: {', '.join(missing_ids)}")

    csl_cache = {
        ref: fetch_csl_item(group_id, str(ref_to_item[ref]["item_key"]), args.api_base)
        for ref in sorted(expected_refs)
    }

    with zipfile.ZipFile(args.source_docx, "r") as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    replaced = replace_placeholders(root, positions, group_id, ref_to_item, item_ids, csl_cache)
    if replaced != len(positions):
        raise SystemExit(f"Expected to replace {len(positions)} placeholders, replaced {replaced}")

    xml_bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    write_docx_with_document_xml(args.source_docx, args.output_docx, xml_bytes)
    print(f"generated={args.output_docx}")
    print(f"citation_fields={replaced}")


if __name__ == "__main__":
    main()
