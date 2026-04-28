from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import httpx


ZOTERO_GROUP_ID_ENV = "ZOTERO_GROUP_ID"
DEFAULT_GROUP_ID = int(os.environ[ZOTERO_GROUP_ID_ENV]) if os.environ.get(ZOTERO_GROUP_ID_ENV) else None
DEFAULT_LOCAL_API_BASE = "http://127.0.0.1:23119"
DEFAULT_WEB_API_BASE = "https://api.zotero.org"


def require_group_id(group_id: int | None) -> int:
    if group_id is None:
        raise SystemExit(f"Provide --group-id or set {ZOTERO_GROUP_ID_ENV}.")
    return int(group_id)


INITIALS_RE = re.compile(r"^[A-Z](?:\s*[A-Z])*$")
ENGLISH_PERSON_RE = re.compile(r"^(?P<family>[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’.-]+)\s+(?P<given>[A-Z](?:\s*[A-Z])*)$")


@dataclass
class Creator:
    creator_type: str = "author"
    family: str | None = None
    given: str | None = None
    literal: str | None = None

    def zotero(self) -> dict[str, str]:
        if self.literal:
            if self.literal == "others":
                return {"creatorType": self.creator_type, "fieldMode": 1, "lastName": "others"}
            return {"creatorType": self.creator_type, "name": self.literal}
        return {
            "creatorType": self.creator_type,
            "firstName": self.given or "",
            "lastName": self.family or "",
        }

    def csl(self) -> dict[str, str]:
        if self.literal:
            return {"literal": self.literal}
        item: dict[str, str] = {}
        if self.family:
            item["family"] = self.family
        if self.given:
            item["given"] = self.given
        return item

    def bibtex(self) -> str:
        if self.literal:
            if self.literal == "others":
                return "others"
            if is_corporate_author(self.literal):
                return "{{" + self.literal + "}}"
            return self.literal
        if self.family and self.given:
            return f"{self.family}, {self.given}"
        return self.family or self.given or ""


@dataclass
class MetadataRecord:
    title: str | None = None
    creators: list[Creator] = field(default_factory=list)
    year: str | None = None
    date: str | None = None
    doi: str | None = None
    url: str | None = None
    item_type: str | None = None
    container_title: str | None = None
    volume: str | None = None
    issue: str | None = None
    pages: str | None = None
    publisher: str | None = None
    language: str | None = None
    source: str | None = None


def normalize_space(text: object) -> str:
    return " ".join(str(text or "").split())


def is_corporate_author(part: str) -> bool:
    corporate_tokens = [
        "中国",
        "国家",
        "国务院",
        "中央",
        "工业和信息化部",
        "国家能源局",
        "华为",
        "ETSI",
        "ITU",
        "CAICT",
        "IEEE",
        "ACM",
    ]
    return bool(part) and any(token in part for token in corporate_tokens)


def split_author_parts(value: str) -> tuple[list[str], bool]:
    value = normalize_space(value)
    has_others = bool(re.search(r"\bet\s+al\.?$", value, flags=re.IGNORECASE)) or value.endswith("等")
    value = re.sub(r",?\s*et\s+al\.?$", "", value, flags=re.IGNORECASE)
    value = value.removesuffix("等")
    parts = [normalize_space(part) for part in re.split(r"\s*[,，、]\s*", value) if normalize_space(part)]
    return parts, has_others


def creator_from_author_part(part: str) -> Creator:
    part = normalize_space(part)
    if not part:
        return Creator(literal="")
    if is_corporate_author(part) or re.search(r"[\u4e00-\u9fff]", part):
        return Creator(literal=part)
    if "," in part:
        family, given = [normalize_space(piece) for piece in part.split(",", 1)]
        return Creator(family=family, given=given)
    match = ENGLISH_PERSON_RE.match(part)
    if match:
        return Creator(family=match.group("family"), given=normalize_space(match.group("given")))
    return Creator(literal=part)


def creators_from_responsibility(value: str | None, include_others: bool = True) -> list[Creator]:
    if not value:
        return []
    parts, has_others = split_author_parts(value)
    creators = [creator_from_author_part(part) for part in parts]
    if include_others and has_others:
        creators.append(Creator(literal="others"))
    return [creator for creator in creators if creator.family or creator.given or creator.literal]


def creators_to_bibtex(creators: Iterable[Creator]) -> str:
    values = []
    for creator in creators:
        rendered = creator.bibtex()
        if rendered:
            values.append(rendered)
    return " and ".join(values)


def creators_to_csl(creators: Iterable[Creator]) -> list[dict[str, str]]:
    return [creator.csl() for creator in creators if creator.literal != "others" and creator.csl()]


def project_item_type(ref_type: str | None, title: str | None = None) -> str | None:
    ref_type = normalize_space(ref_type).upper()
    title_text = normalize_space(title)
    title_lower = title_text.lower()
    if re.search(r"\b(?:itu-t\s+)?(?:recommendation|q\.\d+|y\.\d+|y\.\d+-series|roadmap)\b", title_lower):
        return "standard"
    if (
        any(token in title_text for token in ("白皮书", "研究报告"))
        or "white paper" in title_lower
        or "research report" in title_lower
    ):
        return "report"
    if ref_type in {"R", "R/OL"}:
        return "report"
    if ref_type in {"S", "S/OL"}:
        return "standard"
    if ref_type == "J":
        return "journalArticle"
    if ref_type == "C":
        return "conferencePaper"
    return None


def project_language(
    title: str | None = None,
    responsibility: str | None = None,
    item_type: str | None = None,
) -> str | None:
    text = f"{normalize_space(title)} {normalize_space(responsibility)}"
    if re.search(r"[\u4e00-\u9fff]", text):
        return "zh-CN"
    if item_type in {"journalArticle", "conferencePaper", "report", "standard"} and re.search(r"[A-Za-z]", text):
        return "en-US"
    return None


def _zotero_creator_to_creator(creator: dict[str, Any]) -> Creator:
    if is_fake_others_creator(creator):
        return Creator(literal="others")
    if creator.get("name"):
        return Creator(literal=normalize_space(creator.get("name")))
    return Creator(
        family=normalize_space(creator.get("lastName") or creator.get("family")) or None,
        given=normalize_space(creator.get("firstName") or creator.get("given")) or None,
    )


def parse_display_authors(
    value: str,
    known_creators: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    known = [
        _zotero_creator_to_creator(creator)
        for creator in (known_creators or [])
        if isinstance(creator, dict)
    ]
    parts, has_others = split_author_parts(value)
    output: list[Creator] = []
    for part in parts:
        parsed = creator_from_author_part(part)
        match = None
        if parsed.family and parsed.given and INITIALS_RE.fullmatch(parsed.given):
            initial = parsed.given[:1]
            for candidate in known:
                if (
                    candidate.family
                    and candidate.given
                    and candidate.family.lower() == parsed.family.lower()
                    and candidate.given[:1].upper() == initial.upper()
                ):
                    match = candidate
                    break
        output.append(match or parsed)
    if has_others:
        output.append(Creator(literal="others"))
    return [creator.zotero() for creator in output]


def map_crossref_work(work: dict[str, Any]) -> dict[str, Any]:
    record = metadata_from_crossref_work(work)
    item: dict[str, Any] = {
        "itemType": record.item_type,
        "title": record.title,
        "creators": [creator.zotero() for creator in record.creators],
        "date": record.date or record.year,
        "DOI": record.doi,
    }
    if record.item_type == "conferencePaper":
        item["proceedingsTitle"] = record.container_title
    else:
        item["publicationTitle"] = record.container_title
    if record.volume:
        item["volume"] = record.volume
    if record.issue:
        item["issue"] = record.issue
    if record.pages:
        item["pages"] = record.pages
    if record.publisher:
        item["publisher"] = record.publisher
    return {key: value for key, value in item.items() if value not in (None, "", [])}


def bibtex_author_value(creators: list[dict[str, Any]]) -> str:
    return creators_to_bibtex(_zotero_creator_to_creator(creator) for creator in creators)


def creators_from_crossref_authors(authors: Iterable[dict[str, Any]]) -> list[Creator]:
    creators: list[Creator] = []
    for author in authors:
        family = normalize_space(author.get("family"))
        given = normalize_space(author.get("given"))
        name = normalize_space(author.get("name"))
        if family or given:
            creators.append(Creator(family=family or None, given=given or None))
        elif name:
            creators.append(Creator(literal=name))
    return creators


def _first(values: object) -> str | None:
    if isinstance(values, list) and values:
        value = normalize_space(values[0])
        return value or None
    value = normalize_space(values)
    return value or None


def _year_from_crossref(message: dict[str, Any]) -> str | None:
    for key in ("published-print", "published-online", "issued"):
        parts = message.get(key, {}).get("date-parts")
        if isinstance(parts, list) and parts and isinstance(parts[0], list) and parts[0]:
            return str(parts[0][0])
    return None


def metadata_from_crossref_work(message: dict[str, Any]) -> MetadataRecord:
    title = _first(message.get("title"))
    container_title = _first(message.get("container-title"))
    doi = normalize_space(message.get("DOI")) or None
    crossref_type = normalize_space(message.get("type"))
    item_type = "conferencePaper" if "proceedings" in crossref_type else "journalArticle"
    year = _year_from_crossref(message)
    return MetadataRecord(
        title=title,
        creators=creators_from_crossref_authors(message.get("author", [])),
        year=year,
        date=year,
        doi=doi,
        url=normalize_space(message.get("URL")) or None,
        item_type=item_type,
        container_title=container_title,
        volume=normalize_space(message.get("volume")) or None,
        issue=normalize_space(message.get("issue")) or None,
        pages=normalize_space(message.get("page")) or None,
        publisher=normalize_space(message.get("publisher")) or None,
        source="crossref",
    )


def metadata_from_openalex_work(work: dict[str, Any]) -> MetadataRecord:
    source = ((work.get("primary_location") or {}).get("source") or {})
    authors = []
    for authorship in work.get("authorships", []):
        display_name = normalize_space((authorship.get("author") or {}).get("display_name"))
        if display_name:
            parts = display_name.split()
            if len(parts) >= 2:
                authors.append(Creator(family=parts[-1], given=" ".join(parts[:-1])))
            else:
                authors.append(Creator(literal=display_name))
    return MetadataRecord(
        title=normalize_space(work.get("display_name")) or None,
        creators=authors,
        year=str(work.get("publication_year")) if work.get("publication_year") else None,
        date=str(work.get("publication_year")) if work.get("publication_year") else None,
        doi=(normalize_space(work.get("doi")) or "").removeprefix("https://doi.org/") or None,
        url=normalize_space((work.get("primary_location") or {}).get("landing_page_url")) or None,
        item_type="conferencePaper" if source.get("type") == "conference" else "journalArticle",
        container_title=normalize_space(source.get("display_name")) or None,
        source="openalex",
    )


def metadata_from_semantic_scholar_paper(paper: dict[str, Any]) -> MetadataRecord:
    creators = []
    for author in paper.get("authors", []):
        name = normalize_space(author.get("name"))
        parts = name.split()
        if len(parts) >= 2:
            creators.append(Creator(family=parts[-1], given=" ".join(parts[:-1])))
        elif name:
            creators.append(Creator(literal=name))
    journal = paper.get("journal") or {}
    return MetadataRecord(
        title=normalize_space(paper.get("title")) or None,
        creators=creators,
        year=str(paper.get("year")) if paper.get("year") else None,
        date=str(paper.get("year")) if paper.get("year") else None,
        doi=(paper.get("externalIds") or {}).get("DOI"),
        item_type="journalArticle",
        container_title=normalize_space(journal.get("name") or paper.get("venue")) or None,
        volume=normalize_space(journal.get("volume")) or None,
        pages=normalize_space(journal.get("pages")) or None,
        source="semantic_scholar",
    )


def load_env_file(path: Path = Path(".env.local")) -> dict[str, str]:
    values: dict[str, str] = {}
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip().strip("'\"")
    return values


def zotero_headers(api_key: str | None = None, version: int | None = None) -> dict[str, str]:
    headers = {"Zotero-API-Version": "3"}
    if api_key:
        headers["Zotero-API-Key"] = api_key
    if version is not None:
        headers["If-Unmodified-Since-Version"] = str(version)
    return headers


def citation_key_from_extra(extra: object) -> str | None:
    if not isinstance(extra, str):
        return None
    for line in extra.splitlines():
        lower = line.lower()
        if lower.startswith("citation key:") or lower.startswith("citation-key:"):
            return line.split(":", 1)[1].strip()
    return None


def zotero_citation_key(data: dict[str, Any]) -> str | None:
    key = data.get("citationKey") or citation_key_from_extra(data.get("extra"))
    return str(key) if key else None


def ensure_citation_key_extra(extra: object, citation_key: str | None) -> str:
    text = str(extra or "").strip()
    if not citation_key:
        return text
    lines = [line for line in text.splitlines() if line.strip()]
    for index, line in enumerate(lines):
        lower = line.lower()
        if lower.startswith("citation key:") or lower.startswith("citation-key:"):
            lines[index] = f"Citation Key: {citation_key}"
            return "\n".join(lines)
    lines.append(f"Citation Key: {citation_key}")
    return "\n".join(lines)


def fetch_zotero_group_items(
    group_id: int | None = DEFAULT_GROUP_ID,
    api_base: str = DEFAULT_LOCAL_API_BASE,
    api_key: str | None = None,
    client: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    group_id = require_group_id(group_id)
    close_client = client is None
    if client is None:
        client = httpx.Client(timeout=30)
    try:
        items: list[dict[str, Any]] = []
        start = 0
        while True:
            response = client.get(
                f"{api_base.rstrip('/')}/api/groups/{group_id}/items",
                params={"format": "json", "include": "data", "limit": 100, "start": start},
                headers=zotero_headers(api_key),
            )
            response.raise_for_status()
            batch = response.json()
            if not batch:
                break
            items.extend(batch)
            start += len(batch)
        return items
    finally:
        if close_client:
            client.close()


def looks_like_inverted_initial_creator(creator: dict[str, Any]) -> bool:
    first = normalize_space(creator.get("firstName"))
    last = normalize_space(creator.get("lastName"))
    if not first or not last:
        return False
    return bool(INITIALS_RE.fullmatch(last) and re.fullmatch(r"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’.-]+", first))


def is_fake_others_creator(creator: dict[str, Any]) -> bool:
    values = [
        normalize_space(creator.get("name")),
        normalize_space(creator.get("firstName")),
        normalize_space(creator.get("lastName")),
    ]
    return any(value.lower() in {"others", "other", "et al", "et al."} for value in values if value)


def _expected_item_type(expected: object) -> str | None:
    if isinstance(expected, MetadataRecord):
        return expected.item_type
    if isinstance(expected, dict):
        value = expected.get("item_type") or expected.get("itemType")
        return str(value) if value else None
    if isinstance(expected, str):
        return expected
    return None


def audit_zotero_metadata_items(
    items: list[dict[str, Any]],
    expected_by_citekey: dict[str, object] | None = None,
) -> dict[str, Any]:
    citekey_counts: dict[str, int] = {}
    rows = []
    expected_by_citekey = expected_by_citekey or {}
    for item in items:
        data = item.get("data", {})
        if not isinstance(data, dict):
            continue
        citation_key = zotero_citation_key(data)
        if citation_key:
            citekey_counts[citation_key] = citekey_counts.get(citation_key, 0) + 1
        creators = data.get("creators") or []
        inverted = [creator for creator in creators if isinstance(creator, dict) and looks_like_inverted_initial_creator(creator)]
        fake_others = [creator for creator in creators if isinstance(creator, dict) and is_fake_others_creator(creator)]
        item_type = data.get("itemType")
        missing_container = bool(
            data.get("DOI")
            and item_type in {"journalArticle", "conferencePaper"}
            and not (data.get("publicationTitle") or data.get("proceedingsTitle"))
        )
        expected_item_type = _expected_item_type(expected_by_citekey.get(citation_key or ""))
        item_type_mismatch = bool(expected_item_type and item_type and expected_item_type != item_type)
        if inverted or fake_others or missing_container or item_type_mismatch:
            rows.append(
                {
                    "citation_key": citation_key,
                    "item_key": data.get("key"),
                    "item_type": item_type,
                    "expected_item_type": expected_item_type,
                    "item_type_mismatch": item_type_mismatch,
                    "title": data.get("title"),
                    "doi": data.get("DOI"),
                    "inverted_initial_creators": len(inverted),
                    "fake_others_creators": len(fake_others),
                    "missing_container": missing_container,
                }
            )
    duplicate_citekeys = sorted(key for key, count in citekey_counts.items() if count > 1)
    return {
        "summary": {
            "items_scanned": len(items),
            "citekeys": len(citekey_counts),
            "duplicate_citekeys": len(duplicate_citekeys),
            "suspicious_items": len(rows),
            "inverted_initial_creator_items": sum(1 for row in rows if row["inverted_initial_creators"]),
            "fake_others_creator_items": sum(1 for row in rows if row["fake_others_creators"]),
            "doi_items_missing_container": sum(1 for row in rows if row["missing_container"]),
            "expected_item_type_mismatches": sum(1 for row in rows if row["item_type_mismatch"]),
        },
        "duplicate_citekeys": duplicate_citekeys,
        "items": rows,
    }


def fetch_crossref_by_doi(doi: str, client: httpx.Client) -> MetadataRecord | None:
    response = client.get(f"https://api.crossref.org/works/{doi}", headers={"User-Agent": "thesis_refs/0.1 (mailto:none@example.com)"})
    if response.status_code == 404:
        return None
    response.raise_for_status()
    return metadata_from_crossref_work(response.json()["message"])


def fetch_openalex_by_doi(doi: str, client: httpx.Client) -> MetadataRecord | None:
    response = client.get(f"https://api.openalex.org/works/doi:{doi}")
    if response.status_code == 404:
        return None
    response.raise_for_status()
    return metadata_from_openalex_work(response.json())


def fetch_semantic_scholar_by_doi(doi: str, client: httpx.Client) -> MetadataRecord | None:
    response = client.get(
        f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}",
        params={"fields": "title,authors,venue,year,externalIds,journal"},
    )
    if response.status_code in {404, 429}:
        return None
    response.raise_for_status()
    return metadata_from_semantic_scholar_paper(response.json())


def enrich_by_doi(doi: str, client: httpx.Client) -> MetadataRecord | None:
    for fetcher in (fetch_crossref_by_doi, fetch_openalex_by_doi, fetch_semantic_scholar_by_doi):
        try:
            record = fetcher(doi, client)
        except httpx.HTTPError:
            record = None
        if record and (record.creators or record.container_title):
            return record
    return None


def build_zotero_update(data: dict[str, Any], record: MetadataRecord) -> dict[str, Any]:
    updated = dict(data)
    citation_key = zotero_citation_key(data)
    if record.item_type:
        updated["itemType"] = record.item_type
    if record.title:
        updated["title"] = record.title
    if record.creators:
        updated["creators"] = [creator.zotero() for creator in record.creators]
    if record.date or record.year:
        updated["date"] = record.date or record.year
    if record.doi:
        updated["DOI"] = record.doi
    if record.url:
        updated["url"] = record.url
    if record.container_title:
        if updated.get("itemType") == "conferencePaper":
            updated["proceedingsTitle"] = record.container_title
            updated.pop("publicationTitle", None)
        else:
            updated["publicationTitle"] = record.container_title
    if record.volume:
        updated["volume"] = record.volume
    if record.issue:
        updated["issue"] = record.issue
    if record.pages:
        updated["pages"] = record.pages
    if record.publisher:
        updated["publisher"] = record.publisher
    if record.language:
        updated["language"] = record.language
    updated["extra"] = ensure_citation_key_extra(updated.get("extra"), citation_key)
    return updated


def changed_fields(original: dict[str, Any], updated: dict[str, Any]) -> list[str]:
    keys = sorted(set(original) | set(updated))
    return [key for key in keys if original.get(key) != updated.get(key)]


def make_patch_entry(item: dict[str, Any], record: MetadataRecord) -> dict[str, Any] | None:
    data = item.get("data", {})
    if not isinstance(data, dict):
        return None
    updated = build_zotero_update(data, record)
    fields = changed_fields(data, updated)
    if not fields:
        return None
    return {
        "citation_key": zotero_citation_key(data),
        "item_key": data.get("key"),
        "version": item.get("version") or data.get("version"),
        "source": record.source,
        "doi": data.get("DOI") or record.doi,
        "changed_fields": fields,
        "before": {field: data.get(field) for field in fields},
        "after": {field: updated.get(field) for field in fields},
        "updated_data": updated,
    }


def write_metadata_patch_report(patch: list[dict[str, Any]], output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    json_path = output_dir / "zotero_metadata_patch.json"
    md_path = output_dir / "zotero_metadata_patch.md"
    json_path.write_text(json.dumps({"items": patch}, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = [
        "# Zotero 元数据补全补丁",
        "",
        f"- Items with changes: {len(patch)}",
        "",
        "| Citekey | DOI | Source | Changed fields | Title/Container after |",
        "|---|---|---|---|---|",
    ]
    for item in patch:
        after = item.get("after", {})
        title = after.get("title") or ""
        container = after.get("publicationTitle") or after.get("proceedingsTitle") or ""
        fields = ", ".join(item.get("changed_fields", []))
        lines.append(
            f"| {item.get('citation_key') or ''} | {item.get('doi') or ''} | "
            f"{item.get('source') or ''} | {fields} | {normalize_space(title or container).replace('|', '/')} |"
        )
    md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return json_path, md_path


def apply_patch_to_zotero_web(
    patch: list[dict[str, Any]],
    group_id: int,
    api_key: str,
    api_base: str = DEFAULT_WEB_API_BASE,
) -> list[dict[str, Any]]:
    results = []
    with httpx.Client(timeout=30) as client:
        for entry in patch:
            item_key = entry["item_key"]
            version = entry.get("version")
            response = client.put(
                f"{api_base.rstrip('/')}/groups/{group_id}/items/{item_key}",
                headers={**zotero_headers(api_key, version if isinstance(version, int) else None), "Content-Type": "application/json"},
                content=json.dumps(entry["updated_data"], ensure_ascii=False).encode("utf-8"),
            )
            results.append({"citation_key": entry.get("citation_key"), "item_key": item_key, "status_code": response.status_code})
            response.raise_for_status()
    return results


def api_key_from_env(env_path: Path = Path(".env.local")) -> str | None:
    values = {**load_env_file(env_path), **os.environ}
    return values.get("ZOTERO_API_KEY") or values.get("ZOTERO_API_TOKEN")
