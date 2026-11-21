#!/usr/bin/env python3
"""
resolvers.py — 把文献标识符（DOI / PMID / arXiv / ISBN）解析成 Zotero 条目 JSON。

零第三方依赖（仅标准库 urllib + json + xml），可独立运行：
    python resolvers.py "10.1038/s41586-020-2649-2"
    python resolvers.py "pmid:32214244"
    python resolvers.py "arxiv:1706.03762"

返回的是 Zotero Web/Local API 写请求使用的扁平条目结构（itemType + 顶层字段）。
"""
from __future__ import annotations

import json
import re
import sys
import urllib.parse
import urllib.request
from typing import Optional

USER_AGENT = "zotero-auto-cite/1.0 (research assistant; local use)"

# Crossref type -> Zotero itemType
_CROSSREF_TYPE_MAP = {
    "journal-article": "journalArticle",
    "proceedings-article": "conferencePaper",
    "book": "book",
    "monograph": "book",
    "book-chapter": "bookSection",
    "book-part": "bookSection",
    "report": "report",
    "dissertation": "thesis",
    "posted-content": "preprint",
    "preprint": "preprint",
    "dataset": "dataset",
    "standard": "standard",
    "peer-review": "journalArticle",
    "component": "document",
    "reference-entry": "document",
    "reference-book": "book",
}


def _get_json(url: str, headers: Optional[dict] = None) -> dict:
    req = urllib.request.Request(url, headers=headers or {"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _get_text(url: str) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def _first(values):
    if not values:
        return None
    if isinstance(values, str):
        return values
    return values[0]


def _date_parts_year(issued: Optional[dict]) -> Optional[str]:
    """Crossref issued -> 'YYYY' 或 'YYYY-MM-DD'。"""
    if not issued:
        return None
    dp = issued.get("date-parts") or []
    if not dp or not dp[0]:
        return None
    parts = [p for p in dp[0] if p is not None]
    if not parts:
        return None
    if len(parts) == 1:
        return str(parts[0])
    if len(parts) == 2:
        return f"{parts[0]}-{int(parts[1]):02d}"
    return f"{parts[0]}-{int(parts[1]):02d}-{int(parts[2]):02d}"


def _crossref_creators(work: dict) -> list[dict]:
    creators = []
    for a in work.get("author") or []:
        fam = (a.get("family") or "").strip()
        giv = (a.get("given") or "").strip()
        if fam:
            creators.append({"creatorType": "author", "firstName": giv, "lastName": fam})
        elif a.get("name"):
            creators.append({"creatorType": "author", "name": a["name"]})
    return creators


def crossref_to_zotero(work: dict, extra: str = "") -> dict:
    """Crossref work (CSL-like) -> Zotero 扁平条目。"""
    ctype = work.get("type") or "journal-article"
    item_type = _CROSSREF_TYPE_MAP.get(ctype, "journalArticle")

    item: dict = {"itemType": item_type}
    title = _first(work.get("title"))
    if title:
        item["title"] = title.strip()
    creators = _crossref_creators(work)
    if creators:
        item["creators"] = creators
    abstract = _first(work.get("abstract"))
    if abstract:
        # Crossref 摘要常带 JATS 标签，去掉后压成单行
        abstract = re.sub(r"<[^>]+>", " ", abstract)
        abstract = re.sub(r"\s+", " ", abstract).strip()
        if abstract:
            item["abstractNote"] = abstract
    doi = work.get("DOI")
    if doi:
        item["DOI"] = doi
    url = work.get("URL")
    if url:
        item["url"] = url
    date = _date_parts_year(work.get("issued") or work.get("published-print") or work.get("published-online"))
    if date:
        item["date"] = date
    publisher = work.get("publisher")
    if publisher:
        item["publisher"] = publisher
    place = _first(work.get("publisher-location"))
    if place:
        item["place"] = place
    pages = work.get("page")
    if pages:
        item["pages"] = pages
    volume = work.get("volume")
    if volume:
        item["volume"] = volume
    issue = work.get("issue")
    if issue:
        item["issue"] = issue

    container = _first(work.get("container-title"))
    if container:
        if item_type == "journalArticle":
            item["publicationTitle"] = container
        elif item_type == "bookSection":
            item["bookTitle"] = container
        elif item_type == "conferencePaper":
            item["proceedingsTitle"] = container
        elif item_type == "book":
            item["series"] = container
        else:
            item["publicationTitle"] = container

    issn = _first(work.get("ISSN"))
    if issn:
        item["ISSN"] = issn
    isbn = _first(work.get("ISBN"))
    if isbn:
        item["ISBN"] = isbn

    # event / conference 信息（conferencePaper 用）
    if item_type == "conferencePaper":
        event = work.get("event") or {}
        conf_name = event.get("name")
        if conf_name:
            item["conferenceName"] = conf_name
        conf_date = _date_parts_year(event.get("start"))
        if conf_date:
            item["date"] = item.get("date") or conf_date

    # 章节信息
    if item_type == "bookSection":
        if container:
            item["bookTitle"] = container

    extra_parts = [e for e in (extra.strip(),) if e]
    if extra_parts:
        item["extra"] = extra_parts[0]
    return item


def resolve_doi(doi: str) -> dict:
    doi = doi.strip().lower().replace("https://doi.org/", "").replace("http://doi.org/", "")
    if doi.startswith("doi:"):
        doi = doi[4:]
    work = _get_json(f"https://api.crossref.org/works/{urllib.parse.quote(doi, safe='')}")["message"]
    item = crossref_to_zotero(work)
    item["DOI"] = doi  # 规范化
    return item


def _parse_pubmed_name(name: str) -> dict:
    """'Smith JA' -> {'lastName':'Smith','firstName':'JA'}"""
    parts = name.split()
    if len(parts) == 1:
        return {"creatorType": "author", "name": parts[0]}
    return {"creatorType": "author", "lastName": parts[0], "firstName": "".join(parts[1:])}


def resolve_pmid(pmid: str) -> dict:
    pmid = pmid.strip().lower().replace("pmid:", "")
    data = _get_json(
        "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
        f"?db=pubmed&id={urllib.parse.quote(pmid)}&retmode=json"
    )
    result = data["result"][pmid]
    item: dict = {"itemType": "journalArticle", "extra": f"PMID: {pmid}"}
    title = result.get("title")
    if title:
        item["title"] = title
    authors = [a.get("name") for a in result.get("authors") or [] if a.get("name")]
    if authors:
        item["creators"] = [_parse_pubmed_name(n) for n in authors]
    journal = result.get("source") or result.get("fulljournalname")
    if journal:
        item["publicationTitle"] = journal
    if result.get("volume"):
        item["volume"] = result["volume"]
    if result.get("issue"):
        item["issue"] = result["issue"]
    if result.get("pages"):
        item["pages"] = result["pages"]
    pubdate = result.get("pubdate") or ""
    m = re.search(r"\b(\d{4})\b", pubdate)
    if m:
        item["date"] = m.group(1)
    # 优先取 DOI
    for aid in result.get("articleids") or []:
        if aid.get("idtype") == "doi" and aid.get("value"):
            item["DOI"] = aid["value"]
            break
    return item


def resolve_arxiv(arxid: str) -> dict:
    arxid = arxid.strip().lower().replace("arxiv:", "")
    xml_text = _get_text(f"http://export.arxiv.org/api/query?id_list={urllib.parse.quote(arxid)}")
    import xml.etree.ElementTree as ET
    ns = {
        "atom": "http://www.w3.org/2005/Atom",
        "arxiv": "http://arxiv.org/schemas/atom",
    }
    root = ET.fromstring(xml_text)
    entry = root.find("atom:entry", ns)
    if entry is None:
        raise ValueError(f"arXiv 未返回条目: {arxid}")
    def _t(tag):
        el = entry.find(f"atom:{tag}", ns)
        return " ".join((el.text or "").split()) if el is not None and el.text else None
    title = _t("title")
    summary = _t("summary")
    published = _t("published")
    abs_url = None
    for link in entry.findall("atom:link", ns):
        if link.get("rel") == "alternate":
            abs_url = link.get("href")
    doi_el = entry.find("arxiv:doi", ns)
    doi = (doi_el.text or "").strip() if doi_el is not None and doi_el.text else None

    creators = []
    for a in entry.findall("atom:author", ns):
        nm = " ".join((a.find("atom:name", ns).text or "").split()) if a.find("atom:name", ns) is not None else None
        if nm:
            creators.append(_parse_pubmed_name(nm))
    item: dict = {
        "itemType": "preprint",
        "repository": "arXiv",
        "extra": f"arXiv: {arxid}",
    }
    if title:
        item["title"] = title
    if creators:
        item["creators"] = creators
    if summary:
        item["abstractNote"] = summary
    if published:
        item["date"] = published[:10]
    if doi:
        item["DOI"] = doi
    if abs_url:
        item["url"] = abs_url
    return item


_IDENTIFIER_PATTERNS = [
    ("doi", re.compile(r"^(?:https?://(?:dx\.)?doi\.org/)?(?:doi:)?(10\.\d{4,9}/[^\s]+)$", re.I)),
    ("pmid", re.compile(r"^pmid:(\d{1,9})$", re.I)),
    ("arxiv", re.compile(r"^arxiv:([a-zA-Z\-\.]+\/\d{7}|\d{4}\.\d{4,5}(?:v\d+)?)$", re.I)),
]


def detect_identifier(text: str) -> tuple[str, str]:
    """识别标识符种类，返回 (kind, value)。无法识别返回 ('', text)。"""
    t = text.strip()
    for kind, pat in _IDENTIFIER_PATTERNS:
        m = pat.match(t)
        if m:
            return kind, t
    return "", t


def resolve(identifier: str) -> dict:
    """入口：识别并解析任意标识符，返回 Zotero 条目 JSON。"""
    kind, value = detect_identifier(identifier)
    if kind == "doi":
        return resolve_doi(value)
    if kind == "pmid":
        return resolve_pmid(value)
    if kind == "arxiv":
        return resolve_arxiv(value)
    raise ValueError(f"无法识别的标识符: {identifier!r}（支持 DOI / pmid: / arxiv:）")


if __name__ == "__main__":
    for arg in sys.argv[1:]:
        print(json.dumps(resolve(arg), ensure_ascii=False, indent=2))
